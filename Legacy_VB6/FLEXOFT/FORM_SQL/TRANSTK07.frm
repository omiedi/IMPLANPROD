VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form TRANSTK07 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Transferencias de Stocks"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   5310
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   1155
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   29
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   5265
      Left            =   4440
      ScaleHeight     =   5235
      ScaleWidth      =   2505
      TabIndex        =   21
      Top             =   0
      Width           =   2535
      Begin VB.CommandButton Command4 
         Caption         =   "Kdx"
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
         Left            =   105
         Picture         =   "TRANSTK07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1800
         Width           =   650
      End
      Begin VB.CommandButton Command26 
         Caption         =   "Setup"
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
         Left            =   105
         Picture         =   "TRANSTK07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   3360
         Width           =   650
      End
      Begin VB.CommandButton Command27 
         Caption         =   "Direct"
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
         Left            =   105
         Picture         =   "TRANSTK07.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   2520
         Width           =   650
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
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
         Left            =   105
         Picture         =   "TRANSTK07.frx":0A56
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   120
         Width           =   650
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Help"
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
         Left            =   105
         Picture         =   "TRANSTK07.frx":0BA0
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   840
         Width           =   650
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00CDDADA&
      Caption         =   "OPERACIONES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   4170
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   570
         ScaleHeight     =   615
         ScaleWidth      =   3105
         TabIndex        =   12
         Top             =   420
         Width           =   3135
         Begin VB.CommandButton Command25 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   630
            Left            =   2000
            Picture         =   "TRANSTK07.frx":0EAA
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Anulación de Transferencia"
            Top             =   0
            Width           =   560
         End
         Begin VB.CommandButton Command43 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   2560
            Picture         =   "TRANSTK07.frx":11B4
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Transferencia de Equipos"
            Top             =   0
            Width           =   560
         End
         Begin VB.CommandButton Command3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   0
            Picture         =   "TRANSTK07.frx":12FE
            Style           =   1  'Graphical
            TabIndex        =   0
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Informar Trans- ferencias de Stock"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   590
            TabIndex        =   13
            Top             =   105
            Width           =   1440
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CDDADA&
      Caption         =   "CONSULTAS - LISTADOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3225
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   4170
      Begin VB.Frame Frame4 
         BackColor       =   &H00CDDADA&
         Caption         =   "RANGO DE FECHAS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1755
         Left            =   2100
         TabIndex        =   17
         Top             =   420
         Width           =   1905
         Begin VB.CommandButton BOTSEL1 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   1440
            Picture         =   "TRANSTK07.frx":1608
            TabIndex        =   6
            Top             =   1260
            Width           =   175
         End
         Begin VB.CommandButton BOTSEL 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   1440
            Picture         =   "TRANSTK07.frx":1912
            TabIndex        =   4
            Top             =   600
            Width           =   175
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   360
            TabIndex        =   3
            Text            =   " "
            Top             =   600
            Width           =   1065
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   360
            TabIndex        =   5
            Text            =   " "
            Top             =   1260
            Width           =   1065
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Desde"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   -10
            TabIndex        =   19
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Hasta"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   40
            TabIndex        =   18
            Top             =   1000
            Width           =   750
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00CDDADA&
         Caption         =   "RANGO NUMÉRICO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1755
         Left            =   210
         TabIndex        =   14
         Top             =   420
         Width           =   1800
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
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
            Left            =   360
            TabIndex        =   2
            Top             =   1260
            Width           =   1065
         End
         Begin VB.TextBox Text7 
            Appearance      =   0  'Flat
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
            Left            =   360
            TabIndex        =   1
            Top             =   600
            Width           =   1065
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Hasta"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   40
            TabIndex        =   16
            Top             =   1000
            Width           =   750
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Desde"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   -10
            TabIndex        =   15
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00CDDADA&
         Caption         =   "SALIDA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   210
         TabIndex        =   10
         Top             =   2310
         Width           =   3795
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
            Left            =   3150
            Picture         =   "TRANSTK07.frx":1C1C
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   95
            Width           =   645
         End
         Begin VB.OptionButton Option5 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Pantalla"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   400
            TabIndex        =   7
            Top             =   315
            Value           =   -1  'True
            Width           =   1140
         End
         Begin VB.OptionButton Option4 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Impresión"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   1710
            TabIndex        =   8
            Top             =   315
            Width           =   1275
         End
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "TRANSTK07.frx":1F26
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "TRANSTK07.frx":215A
      TabIndex        =   28
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Menu mnuTransacciones 
      Caption         =   "Transacciones"
      Begin VB.Menu mnuDepNegPañFac 
         Caption         =   "Depurar Negativos Depósito Pañol de Facturación"
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "Ver"
      Begin VB.Menu mnuReimprimeTransferencias 
         Caption         =   "Reimprimir Transferencia de Stock"
      End
   End
End
Attribute VB_Name = "TRANSTK07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RUSANJUA$

Sub CARGARBOLREDET(DEPOORIGEN%)
       BCODECTP.Label4 = DEPOORIGEN%
       BCODECTP.Text1 = ""
       BCODECTP.List1.Clear
       BCODECTP.Show 1
       If BCODECTP.List1.ListCount > 0 Then
        If BCODECTP.VALI_BCODE09 = "APROB" Then
         '2.- CARGAR BOLREDET
         JJ& = 0
         REBLA$ = REGBLAN$("!BOLREDET")
         For JX% = 0 To BCODECTP.List1.ListCount - 1
           
           Z$ = BCODECTP.List1.List(JX%)
           CI1$ = Mid$(Z$, 1, 17)
           CI2% = CODINT%(CI1$)
           Q1# = XVALO(Mid$(Z$, 49, 8))
           LOTEX$ = Mid$(Z$, 61, 12)
           
           Y$ = REBLA$
           Call REPLA(Y$, MKI$(CI2%), 83, 2)
           Call REPLA(Y$, Left$(Unimed$(CI2%), 4), 85, 4)
           Call REPLA(Y$, MKD$(Q1#), 89, 8)
           Call REPLA(Y$, LOTEX$, 69, 12)
           Call REPLA(Y$, LOTEX$, 107, 12)
           JJ& = JJ& + 1
           Call GRAREG("!BOLREDET", Y$, JJ&)
         Next JX%
         Call SETULTREG("!BOLREDET", JJ&)
       ElseIf BCODECTP.List1.ListCount <= 0 Then
         Call SETULTREG("!BOLREDET", 0)
       End If

  End If
End Sub

Private Sub BOTSEL_Click(Index As Integer)
    Call SELECHO("SELFECHA")
    Text2.TEXT = VALACT1$("SELFECHA")
End Sub

Private Sub BOTSEL1_Click(Index As Integer)
    Call SELECHO("SELFECHA")
    Text1.TEXT = VALACT1$("SELFECHA")
End Sub

Private Sub Command1_Click()
    Hide
    Call FRESHALL
End Sub

Private Sub Command10_Click()
End Sub

Private Sub Command2_Click()
    DESNUM& = 1: If XVALO(Text7.TEXT) > 0 Then DESNUM& = XVALO(Text7.TEXT)
    DESNUX$ = "TR." + TRIM$(Str$(DESNUM&))
    HASNUM& = 999999
    If TRIM$(Text8.TEXT) <> "" Then
        If XVALO(Text8.TEXT) >= DESNUM& Then HASNUM& = XVALO(Text8.TEXT)
    End If
    HASNUX$ = "TR." + TRIM$(Str$(HASNUM&))
    DESFE% = 0: If FECHANUM(TRIM$(Text2.TEXT)) > 0 Then DESFE% = FECHANUM(TRIM$(Text2.TEXT))
    HASFE% = 32767
    If TRIM$(Text1.TEXT) <> "" Then
        If FECHANUM(TRIM$(Text1.TEXT)) >= DESFE% Then HASFE% = FECHANUM(TRIM$(Text1.TEXT))
    End If
    '
    Screen.MousePointer = 11
    Call SETULTREG("LISTRAJ", 0)
    '
    FEC1 = FETEXCOR1$(DESFE%)
    FEC2 = FETEXCOR1$(HASFE%)
    '
    Call ABRESTOCKS
    SQLX$ = " SELECT ReferCor, COD_INTE, DEPOMOVI, CANTINGRE, DoPriCor, FECHMOV, CANTSALID FROM MOVISTO WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE CodiMovi =  'TR' "
    SQLX$ = SQLX$ + " and DoPriCor >= '" & DESNUX$ & "' "
    SQLX$ = SQLX$ + " AND DoPriCor <= '" & HASNUX$ & "' "
    SQLX$ = SQLX$ + " AND FechMov >= '" & FEC1 & "' "
    SQLX$ = SQLX$ + " AND FechMov <= '" & FEC2 & "' "
    'SQLX$ = SQLX$ + "ORDER BY VAL(MID$(DoPriCor,4,7)) ASC"
    SQLX$ = SQLX$ + "ORDER BY CAST(SUBSTRING(DoPriCor,4,7)AS INTEGER) ASC"
    '
'    Dim AJUSMOVTMP As ADODB.Recordset
'    Set AJUSMOVTMP = New ADODB.Recordset
'    AJUSMOVTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
    '
    Set AJUSMOVTMP = READSET(SQLX$)
    With AJUSMOVTMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        Call MENSERR(24, "No Existen Transferencias Registradas segun Filtro")
        On Error GoTo 0
        Screen.MousePointer = 1
        Exit Sub
      End If
      On Error GoTo 0
      Do While .EOF = False
      REFE$ = SAFETEXT$(!ReferCor)
      CI1% = XVALO(!cod_inte)
      DP1% = XVALO(!depomovi)
      CAN# = XVALO(!CANTINGRE)
      FF% = CVINT(!FechMov)
      NRO& = Mid(SAFETEXT$(!dopricor), 4)
      Y$ = REGBLAN$("LISTRAJ")
      Call REPLA(Y$, MKI$(FF%), 1, 2)
      Call REPLA(Y$, REFE$, 3, 48)
      Call REPLA(Y$, MKS$(NRO&), 51, 4)
      Call REPLA(Y$, MKI$(CI1%), 55, 2)
      Call REPLA(Y$, Unimed$(CI1%), 57, 4)
      If CAN# > 0 Then
          Call REPLA(Y$, MKD$(CAN#), 61, 8)
          Call REPLA(Y$, MKD$(!CANTSALID), 69, 8)
        Else
          Call REPLA(Y$, MKD$(CAN#), 61, 8)
          Call REPLA(Y$, MKD$(!CANTSALID), 69, 8)
      End If
      Call REPLA(Y$, Chr$(DP1%), 80, 1)
      Call REGAPP("LISTRAJ", Y$)
      
      .MoveNext
9      Loop
    End With
    '
    Call CIERRARCH("LISTRAJ")
    Screen.MousePointer = 1
    '
    OPCI% = 1
    If Option4.Value = True Then OPCI% = 2
    If OPCI% = 1 Then
        'Call OPNOIN("")
        Call COMUNI("Solo Disponible para Impresión. Genere el \ Listado y Exportelo a un Archivo de Texto")
        GoTo 10
      ElseIf OPCI% = 2 Then
10      Call CONECRUN("*LITRAST", "")
    End If
    '
End Sub

Private Sub Command25_Click()
                Command25.Enabled = False
                TituloLista$ = "Anulación de Transferencias"
                SQLX$ = "SELECT  DOPRILAR,DOSECLAR,REFERCOR, FECHMOV,COD_INTE, COD_EXTE,DESCRIP, CANTINGRE,CANTSALID "
                SQLX$ = SQLX$ + " FROM  MOVISTO WHERE CODIMOVI = 'TR' ORDER BY FEREMOVI,DOSECLAR"
  
  
                Campos$ = "ID Transf/A10;DOSECLAR/A0;Referencia/A24;Fecha/D8;CI/A0;Código/A18;Descripción/A36;C.Salida/F12.1;C.Ingresada/F12.1"
                Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
                Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
                FRMZELECHO.Caption = TituloLista$
    
                'para agrupar la columna 1 donde esta el id de transferencia
                FRMZELECHO.MSF2.FixedCols = 1
                FRMZELECHO.MSF2.ScrollTrack = True
                FRMZELECHO.MSF2.ScrollBars = flexScrollBarBoth
                FRMZELECHO.MSF2.MergeCells = flexMergeFree
                FRMZELECHO.MSF2.MergeCol(1) = True
                '*********************************************************

                Set Rst = FLEXCONN.Execute(SQLX$)
                Call Carga_MSF_Recordset(Rst, Campos$, FRMZELECHO.MSF2, 1)
                On Error Resume Next
                Rst.Close
                Set Rst = Nothing
                On Error GoTo 0
                    
                FRMZELECHO.Width = 16500
406             FRMZELECHO.Show 1
                
                IDTRANSF$ = TRIM(RESP_ZELE_VECT(1)) 'TR.1
                IDDOSECLAR$ = TRIM(RESP_ZELE_VECT(2)) 'TR.10305
                If IDTRANSF$ = "" Then GoTo 99
                OPX% = COMALTER%("¿Esta Seguro que Desea Anular la Transferencia '" & IDTRANSF$ & "'?\\No, Cancelar\Si, Anular")
                If OPX% = 2 Then
                  SQLX$ = "DELETE MOVISTO WHERE DOPRILAR = '" & IDTRANSF$ & "'"
                  SQLX$ = SQLX$ + " AND DOSECLAR = '" & IDDOSECLAR$ & "'"
                  FLEXCONN.Execute (SQLX$)
                End If
99              Command25.Enabled = True
                
End Sub



Private Sub Command26_Click()
   Call CONECRUN("FSETUP04/OPTRAJUST", "Configuración de Funcionamiento")
   Call FINAL("")
   'OPMOMAQ04.Show 1
End Sub

Private Sub Command27_Click()
    Command27.Enabled = False
10  OPX% = ALTERNA%("Aprobación de Transferencias\Pendientes de Aprobación\Anulación de Transferencias Pendientes\Otros Accesos Directos")
    If OPX% = 1 Then
        Call TRANSFE_APRO(0)
        GoTo 10
    ElseIf OPX% = 2 Then
        Call TRANSFE_APRO(1)
        GoTo 10
    ElseIf OPX% = 3 Then
        Call ANULATRANSFEPEN
        GoTo 10
    ElseIf OPX% = 4 Then
         Call ACCDIR("TRANSFE")
    End If
    Command27.Enabled = True
End Sub

Private Sub Command3_Click()
    Call TRASTOCK
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    Call SITSTOK07.Command3_Click
    'Call CONECRUN("*KARDEX", "")
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Call OPNOIN("")
End Sub

Private Sub Command43_Click()
   TRANEQUI07.Show 1
End Sub

Private Sub Command6_Click()
    Call HELPONLINE("TRANSFE")
End Sub

Private Sub Form_Load()
   '
   Call TRANSLABELS(Name)

   Call APLICACIONSKINS(Me, Picture1)
   '
   If EXISTE%(LUDAT$ + "TIPOTRAN.RFS") > 0 Then
    Call PoneEnCeroNull("MOVISTO", "TIPOTRAN", REGAF&)
    Call GRAREG("TIPOTRAN", Chr$(0) + "Terminado y Envasado", 1)
    Call GRAREG("TIPOTRAN", Chr$(1) + "Recambio", 2)
    Call GRAREG("TIPOTRAN", Chr$(2) + "No Conforme", 3)
    Call CIERRARCH("TIPOTRAN")
   End If
   '
   Option5.Value = True
   '
End Sub

Function CMOVDAT_TRAN&()
   
   NUME_MAX& = 0
   FIN& = ULTREG&("APROTRAN")
   For i& = 1 To FIN&
      TX$ = REGLEIDO("APROTRAN", i&)
      DOPRI$ = Mid$(TX$, 23, 10) 'AJ.000001'
      NUME& = XVALO(Mid$(DOPRI$, 4)) '1
      If NUME_MAX& < NUME& Then
         NUME_MAX& = NUME&
      End If
   Next i&
   '
   CMOVDAT_TRAN& = NUME_MAX&
   '
End Function

Private Sub mnuDepNegPañFac_Click()
    '
    '
    DEPOFACTU% = XVALO(Control$("", "DEPOEXPE"))
    PANOLFACTU% = XVALO(Control$("", "DPREPPAÑ"))
    DEPOPAÑOL1% = XVALO(Control$("", "DEPOCEN"))
    
20  Campos$ = "Ci/A0;Código/A16;Descripción/A32;Depósito: " & PANOLFACTU% & "/F14.1;Depósito: " & DEPOPAÑOL1% & "/F14.1;Cant.a Transf./F14.1"
    
    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
    Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)
    
    
    GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
    GRILLAGRAL.mnuMenu1.Visible = True

    GRILLAGRAL.mnuMenu2.Caption = "Depurar Negativos Seleccionados"
    GRILLAGRAL.mnuMenu2.Visible = True
    GRILLAGRAL.mnuMenu3.Caption = "Seleccionar/Revertir Selección"
    GRILLAGRAL.mnuMenu3.Visible = True

    GRILLAGRAL.GRID.ZOrder 0
    GRILLAGRAL.TXTBUSCAR = ""

    
    'TOMA LA FECHA DE HOY Y LE RESTO LOS DIAS DE TOLERANCIA (CANTIDIAS%)
    FEHOY% = HOY(HO$)
    
    FECHOYTOLERA = FETEXCOR1$(FEHOY%) 'FECHA DE HOY - LOS DIAS DE TOLERANCIA
    '
    'CARGO EL RECORDSET
    Dim obj As New RECORXET
    Set RS1 = obj.RS_NEGATIVOS_MDT(PANOLFACTU%)
    Call Carga_MSF_Recordset(RS1, Campos$, GRILLAGRAL.GRID, 1)
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    '*******************************************
    Screen.MousePointer = 1
    If GRILLAGRAL.GRID.Rows < 2 Then Exit Sub
    
    TX$ = "Reporte de Negativos del Depósito: " & PANOLFACTU% & " (Solo Items No Trazables)"
    GRILLAGRAL.Caption = TX$
    '*******************************************
    'AHORA RECORRO LA GRILLA Y LE ASIGNO EL STOCK DEL 50 A CADA UNO Y TAMBIEN AGREGO LA CANTIDAD QUE SE PODRIA TRANSFERIR
    
    For i& = 1 To GRILLAGRAL.GRID.Rows - 1
       CI1% = XVALO(GRILLAGRAL.GRID.TextMatrix(i&, 1))
       CANTSTOCK = STODEPOS(CI1%, DEPOPAÑOL1%)
       If CANTSTOCK <= 0 Then CANTSTOCK = 0
       If CANTSTOCK > Abs(XVALO(GRILLAGRAL.GRID.TextMatrix(i&, 4))) Then
          CANTATRANSFERIR = Abs(XVALO(GRILLAGRAL.GRID.TextMatrix(i&, 4)))
       Else
          CANTATRANSFERIR = CANTSTOCK
       End If
       GRILLAGRAL.GRID.TextMatrix(i&, 5) = FORMATNUM$(CANTSTOCK, "F12.1")
       GRILLAGRAL.GRID.TextMatrix(i&, 6) = FORMATNUM$(CANTATRANSFERIR, "F12.1")
       
    Next i&
    
30  GRILLAGRAL.Show 1
    Item1% = 0
    FF% = HOY(HO$)
    Dim MSF As msFlexGrid
    Set MSF = GRILLAGRAL.GRID
    If GRILLAGRAL.TXTBUSCAR <> "" Then Set MSF = GRILLAGRAL.MSF_2
    If GRILLAGRAL.MENU2 > 0 Then
            Call COMUNI("Atención!!! \Solo Se Depuraran los Códigos con Cantidades Negativas\Que contengan Cantidad (Suficiente o Parcial) en el Depósito: " & DEPOPAÑOL1%)
            For i& = 1 To MSF.Rows - 1
               If MSF.TextMatrix(i&, 0) = "*" Then
                 Item1% = Item1% + 1
                 CI1% = XVALO(MSF.TextMatrix(i&, 1))
                 If PRONUMTRANSF& < 1 Then   'PARA QUE SOLO ENTRE UNA VEZ
                   CMOV$ = TRIM$(Control$("", "CODTRAST"))
                   If CMOV$ = "" Then CMOV$ = "TR"
                   PRONUMTRANSF& = ProNroComprobante("TR")
                   DOPRI$ = "TR." + TRIM$(Str$(PRONUMTRANSF&))
                 End If
                 CANTAA = XVALO(MSF.TextMatrix(i&, 6))
                 If CANTAA > 0 Then  'si tiene cantidad para transferir transfiere
                   Call MOVISTOK(FF%, CI1%, "", CMOV$, DOPRI$, DOPRI$, DOPRI$, DOPRI$, "Sal.Aut.Depuración " & DEPOPAÑOL1%, NORIT%, 0, 1, 0, DEPOPAÑOL1%, -1 * CANTAA, 0, , "", , , 0)     '(-1) * CANTAA )
                   Call MOVISTOK(FF%, CI1%, "", CMOV$, DOPRI$, DOPRI$, DOPRI$, DOPRI$, "Ing.Aut.Depuración " & PANOLFACTU%, NORIT%, 0, 1, 0, PANOLFACTU%, 1 * CANTAA, 0, , "", , , 0)   '(-1) * CANTAA )
                 End If
               End If
            Next i&
            GRILLAGRAL.MENU2 = 0
            Screen.MousePointer = 1
            If Item1% < 1 Then
               Call COMUNI("No se Han Seleccionado Items a Depurar"): GoTo 30
            Else
               Call COMUNI("Proceso Realizado")
            End If
     ElseIf GRILLAGRAL.MENU1 > 0 Then
            Set MSF = GRILLAGRAL.GRID
            If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
            Call CARGA_TABLA_IMPRE(MSF, Campos$, "DEPUNEGA", "Depuración de Items Negativos en Depósito: " & PANOLFACTU%, "TABLA_IMPRE", RET_NAMECONS$, 1)
            Call FINAL("?" & RET_NAMECONS$)
            GRILLAGRAL.MENU1 = 0
            Screen.MousePointer = 1
     End If

     
     
'    If GRILLAGRAL.MENU1 > 0 Then
'       If DERACCE%("DEPOCVE", "Depuración de Ordenes de Compra Vencidas") >= 2 Then
'            Set MSF = GRILLAGRAL.GRID
'            If TRIM$(GRILLAGRAL.TXTBUSCAR) <> "" Then Set MSF = GRILLAGRAL.MSF_2
'            For i& = 1 To MSF.Rows - 1
'               If MSF.TextMatrix(i&, 0) = "*" Then
'                 ITEMANUL% = ITEMANUL% + 1
'                 NOCOMPRA& = XVALO(MSF.TextMatrix(i&, 1))
'                 NP% = XVALO(MSF.TextMatrix(i&, 2))
'                 CIXI% = CODINT%(MSF.TextMatrix(i&, 4))
'                 NUOR% = XVALO(MSF.TextMatrix(i&, 11))
'                 CONDI$ = " NUME_OCOM = " & NOCOMPRA& & " AND COD_INTE = " & CIXI%
'                 CONDI$ = CONDI$ + " AND NPRO_OCOM = " & NP%
'                 CONDI$ = CONDI$ + " AND NUORD_ITEM = " & NUOR%
'                 CONDI$ = CONDI$ + " AND Stat_Ocom < 8"
'                 '
'                 Call ACTUREGISTRO("OCOMDETA", "Stat_Ocom", CONDI$, 8, REGAF&, "NOMESS")
'                 GRILLAGRAL.MENU1 = 0
'               End If
'            Next i&
'            '
'            If ITEMANUL% < 1 Then
'               Call COMUNI("No se Ha Seleccionado Ningún Registro")
'               GRILLAGRAL.MENU1 = 0 ' lo vuelvo a poner en 0 la bandera para saber si se cliqueo el menu1
'               GoTo 30 'presento la ventana nuevamente
'            Else
'               Call COMUNI("Se Han Anulado " & ITEMANUL% & " Items Seleccionados")
'               GoTo 20
'            End If
'       End If
'    ElseIf GRILLAGRAL.MENU2 > 0 Then  'ENVIO DE MAIL
'            GRILLAGRAL.MENU2 = 0
'            Set MSF = GRILLAGRAL.GRID
'            If TRIM$(GRILLAGRAL.TXTBUSCAR) <> "" Then Set MSF = GRILLAGRAL.MSF_2
'            Call CARGA_ENCABEZADO(MAILNEW.MSF, CAMPOS$, MAILNEW)
'            Call CopiarMsfAOtroMsf(MSF, MAILNEW.MSF, 1, 0, MAILNEW)
'
'            MAILNEW.Show 1
'    End If
End Sub


Private Sub mnuReimprimeTransferencias_Click()
    
    Campos$ = "Fecha/D8;Transferencia/A16s"
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    
    Dim Rst As ADODB.Recordset
    ConsultaSql$ = "select fechmov, dopricor from movisto with(nolock) where codimovi='TR' and dopricor like 'TR.%' group by fechmov, dopricor order by fechmov desc"
    Set Rst = READSET(ConsultaSql$)
    
    Fila& = 0
    FRMZELECHO.MSF2.Rows = Rst.RecordCount + 1
    With Rst
        Do While Not .EOF
            FechaTransferencia% = CVINT(!FechMov)
            DOCUMENTO$ = !dopricor
                        
            Fila& = Fila& + 1
            FRMZELECHO.MSF2.TextMatrix(Fila&, 1) = FECHATEX$(FechaTransferencia%)
            FRMZELECHO.MSF2.TextMatrix(Fila&, 2) = DOCUMENTO$
            
            .MoveNext
        Loop
    End With
    
    On Error Resume Next
    Rst.Close
    Set Rst = Nothing
    On Error GoTo 0
    
    FRMZELECHO.Width = 5000
    FRMZELECHO.Caption = "Transferencias de Stock"
15  FRMZELECHO.Show 1

    DOCUMENTO$ = RESP_ZELE_VECT(2)
    If DOCUMENTO$ <> "" Then
        NumeroBusqueda$ = TRIM$(Mid$(DOCUMENTO$, 4))
        NumeroCompleto$ = NumeroBusqueda$
        DOCUNU& = DOCID("VALE TRANSFERENCIA DE STOCK", NumeroBusqueda$, NumeroCompleto$)
        Call MUESTRADOC(DOCUNU&)
    
        GoTo 15
    End If
End Sub

'
Private Sub Text1_DblClick()
    Call SELECHO("SELFECHA")
    Text1.TEXT = VALACT1$("SELFECHA")
End Sub

Private Sub Text2_DblClick()
    Call SELECHO("SELFECHA")
    Text2.TEXT = VALACT1$("SELFECHA")
End Sub

Sub TRASTOCK()
    '
    EPAC$ = UCase$(TRIM$(EMPREAC$))

    MODOALEJPA% = 0
    HOOI% = HOY(HO$)
    FORIPRE$ = TRIM$(Control$("", "FORTRAST"))
    If FORIPRE$ <> "" Then
        LU$ = LUDAT$
        If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
            RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
            Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Vale de Transferencia:\  \'" + RMCC$ + "'.")
            Exit Sub
        End If
        '
    End If
    '
100 CMOV$ = TRIM$(Control$("", "CODTRAST"))
    If CMOV$ = "" Then CMOV$ = "TR"
    '
    Call ABRESTOCKS
    'sqlx$ = " SELECT * FROM MOVISTO "
    'sqlx$ = sqlx$ + "WHERE CodiMovi = '" & CMOV$ & "'"
    'sqlx$ = sqlx$ + "ORDER BY VAL(MID$(DoPriCor,4)) DESC "
    '
    '///OT-8-813-05/09/08
    SQLX$ = " SELECT max(CAST(substring(DoPriCor, 4,12) AS INTEGER)) as Maximo FROM MOVISTO"
    SQLX$ = SQLX$ + " WHERE CodiMovi = '" & CMOV$ & "'"
    ' OT-8-870-24/10/08 - SE INCLUYE CONDICION PORQUE SE ENCONTRO EN GUIDI UNA DEVOLUCION
    ' CON EL TIPO DE MOVIMIENTO TR Y EL NUMERO DE COMPROBANTE 1-22687 EL CUAL NO PODIA CONVERTIR A INTEGER
    SQLX$ = SQLX$ + " AND SUBSTRING(DOPRICOR,1,2) <> 'RT'"
    ' OT-8-870-24/10/08-FIN
    '
    '///OT-8-813-FIN
    UNUM& = 0
    'Set MoviTemp = Stocks.OpenRecordset(sqlx$, dbOpenSnapshot)
    'On Error Resume Next
    'MoviTemp.MoveFirst
    'If Err < 1 Then
      'UNUM& = XVALO(Mid$(MoviTemp!DoPriCor, 4))
    'End If
    Dim MIRS As ADODB.Recordset
    Set MIRS = FLEXCONN.Execute(FILTSQL$(SQLX$))
    If Not IsNull(MIRS!Maximo) Then
        UNUM& = XVALO(MIRS!Maximo)
    End If
    
    NRO& = 1 + UNUM&
    MIRS.Close
    Set MIRS = Nothing
    
    '
    TIPODOC$ = "Vale Transferencia de Stock"
    NUMEDOC$ = TRIM$(Str$(NRO&))
    '
110 Screen.MousePointer = 1
    '
    VDEF$ = String$(6, 0) + MKI$(HOOI%) + Space$(16) + MKS$(0) + Space$(32)
    OPX% = COMALTER%("Opciones de Procesamiento:\\Item único\Lista Múltiple\Vaciamiento de Deposito\Transferencia y Asignación x Pedidos\Cancelar")
    If OPX% < 1 Or OPX% > 4 Then Exit Sub
    '
    If OPX% = 3 Then ' DERECHO DE ACCESO SOLO PARA VACIAMIENTO DE DEPOSITO
       If DERACCE%("VACIADEP", "Vaciamiento de Depósito") < 2 Then
         GoTo 110
       End If
    End If
    
    If OPX% = 4 Then ' Transfencia y Asignación  a Pedidos
       ASINUMSE.Show 1
       GoTo 110
    End If
    
   If OPX% = 2 Or OPX% = 3 Then
      '\\\OT-5-0447-WAR-08/07/05///
      FORIPRE$ = TRIM$(Control$("", "FORTRSTM"))
      If FORIPRE$ <> "" Then
        If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") < 1 Then
          FORIPRE$ = ""
        End If
      End If
      '
      'If FORIPRE$ = "" Then
      '  Call COMUNI("Error de Configuración:\  \Falta Definir Formulario\de Impresión de Vale de Transferencia.")
      '  Exit Sub
      'End If
      '\\\OT-5-0447-WAR-FIN///
      
115   If InStr(EMPREAC$, "ALEJPA") < 1 Then
          X$ = OBJPLA$("TRANSTOMUL", VDEF$)
      Else
          MODOALEJPA% = 1
          'OBJETO IGUAL PERO CON TIPO DE TRANSFERENCIA
          X$ = OBJPLA$("TRANTIPO", VDEF$)
      End If
      If X$ = "" Then
        Exit Sub                        ' abandonar
      End If
      '
      VDEF$ = X$
      DepOrig% = Asc(Mid$(X$, 3, 1)): DEPOSI$ = ECOARCH$("ECODEP", Chr$(DepOrig%))
      DepoDes% = Asc(Mid$(X$, 6, 1)): DEPODE$ = ECOARCH$("ECODEP", Chr$(DepoDes%))
      FEMOVI% = CVI(Mid$(X$, 7, 2))
      REFMO$ = TRIM$(Mid$(X$, 29, 32))
      If REFMO$ = "" Then
        Call MENSERR(24, "Referencia o Motivo es Mandatoria")
        GoTo 115
      End If
      '
      TIPOTRA% = 0
      If MODOALEJPA% > 0 Then TIPOTRA% = Asc(Mid$(X$, 61, 1)) 'LEO EL TIPO DE TRANSFERENCIA
      
      'VALIDACION DE FECHA DE TRANSFERENCIA SEGUN CONTROL DEL SETUP
      If DERACCE%("CAMFETRA/NOMESS", "Validación de Fecha de Transferencia") < 2 Then
        If Control$("STOCKS", "VALFETRA") = "SI" Then
          DIAHO% = HOY(HO$)
          If FEMOVI% <> DIAHO% Then
             Call COMUNI("No Válido Realizar Transferencia Con Fecha Diferente a la de Hoy")
             GoTo 115
          End If
        End If
      End If
      '
      If InStr(EMPREAC$, "ALEJPA") > 0 Then ' SOLO ALEJPA
116     OPXBARRA% = COMALTER%("Opciones de Carga\\Carga Manual\Código de Barras")
        If OPXBARRA% < 1 Or OPXBARRA% > 2 Then GoTo 110
        '
        If OPXBARRA% = 2 Then
           LECCODBA.List1.Clear
117        LECCODBA.Show 1
           '
           If XVALO(LECCODBA.List1.ListCount) < 1 Then GoTo 116
           '
           Call SETULTREG&("!BOLREDET", 0)
           For U& = 1 To LECCODBA.List1.ListCount
           CI% = LECCODBA.LECT_CODIGO%(LECCODBA.List1, U& - 1)
           CANTI = LECCODBA.LECT_CANTIDAD(LECCODBA.List1, U& - 1)
            '
            Y$ = REGBLAN$("!BOLREDET")
            Call REPLA(Y$, MKI$(CI%), 83, 2)
            Call REPLA(Y$, Unimed$(CI%), 85, 4)
            Call REPLA(Y$, MKD$(CDbl(CANTI)), 89, 8)
            Call REPLA(Y$, Chr$(XVALO(U&)), 98, 1)
            Call REPLA(Y$, MKD$(CDbl(CANTI)), 163, 8)
            '
            Call GRAREG("!BOLREDET", Y$, U&)
           Next U&
           '
'           LECCODBA.List1.Clear
           JJ& = ULTREG&("!BOLREDET")
           GoTo 1110
        End If
     End If
      
1112  Call SETULTREG("!BOLREDET", 0)
      If InStr(1, EMPREAC$, "ARCON") > 0 Then
        OPX% = COMALTER%("Opciones Tipo Transferencia:\\Modo Tradicional\Asignado a Pedidos")
        If OPX% < 1 Then GoTo 9999
        If OPX% = 2 Then
           Screen.MousePointer = 1
           Call ValeMatXPeDI
        End If
        
     End If

      Call CIERRARCH("*.*")
      '
      Screen.MousePointer = 11
      '
      If OPX% = 3 Then
        JJ& = 0
        For i% = 1 To NUMAS%
          CIXI% = i%
          If ESTRAZABLE%(CIXI%) > 0 Then GoTo 10
          STO# = STODEPFECH(CIXI%, DepOrig%, FEMOVI%)
          If STO# > 0.05 Then
               XX$ = REGBLAN$("BOLREDET")
               Call REPLA(XX$, MKI$(CIXI%), 83, 2)
               Call REPLA(XX$, MKD$(STO#), 89, 8)
               JJ& = JJ& + 1
               Call GRAREG("!BOLREDET", XX$, JJ&)
          End If
10      Next i%
        Call SETULTREG("!BOLREDET", JJ&)
        Call COMUNI("En el Vaciamiento Se Han Excluido los Articulos Trazables ")
      End If
      '
1110  fff$ = REPLACAR$(FECHATEX$(FEMOVI%), "/", "-")
      If InStr(EMPREAC$, "CAMBACERES") > 0 Then
         Screen.MousePointer = 1
         Call CARGARBOLREDET(DepOrig%)
      End If
      
      AVENTA$ = "BOLREDET"
      If EXISTE%("TRAZAB07.EXE") > 0 Or EXISTE%("TRAZAB08.EXE") > 0 Then AVENTA$ = "MOVITRAZ"
1113  VTB% = VENTABU%(AVENTA$ + "/NC/TITUPAN=Transferencia de Stock - De " + TRIM$(DEPOSI$) + " a " + TRIM$(DEPODE$) + " - " + fff$)
      If VTB% < 0 Or ULTREG&("!BOLREDET") < 1 Then
         If OPXBARRA% < 2 Then GoTo 110 ' ESTO COMO ESTABA
         If OPXBARRA% = 2 Then GoTo 117 ' SI VIENE DE CODIGO DE BARRA VUELVE AL FRM DE CODIBARRA
      End If
      '
      CONTA% = ULTREG&("!BOLREDET")
      '
1115  If CONTA% > 0 Then
        Screen.MousePointer = 11  'Call BLIMESS("Registrando Recepcion ...")
        '
        '\\\OT-05-0730-WAR-10/11/05///
        CAITMO% = 0
        For YX& = 1 To ULTREG&("!BOLREDET")
             YY$ = REGLEIDO$("!BOLREDET", YX&)
             CI% = CVI(Mid$(YY$, 83, 2))
             CAN# = CVD(Mid$(YY$, 89, 8))
             If CI% > 0 Then
               If CI% <= NUMAS% Then
                 If CAN# > 0.005 Then
                   CAITMO% = CAITMO% + 1
                   CIMOVIM%(CAITMO%) = CI%
                   CAMOVIM(CAITMO) = (-1) * CAN#
                 End If
               End If
             End If
        Next YX&
        If PUEDEMOVI%(DepOrig%) < 0 Then GoTo 1113 ' CANCELA SALIDA
        '\\\OT-05-0730-WAR-FIN///
        '
        ' VALIDA LOS DATOS INGRESADOS
        
        HUBOSELOT% = 0
        For YX& = 1 To ULTREG&("!BOLREDET")
           YY$ = REGLEIDO$("!BOLREDET", YX&)
           CI% = CVI(Mid$(YY$, 83, 2))
           If ESTRAZABLE%(CI%) > 0 Then
               ' OBLIGA A INGRESAR EL LOTE
               ' COMPARA CANTIDAD DESPACHADA CON SALDO DEL LOTE
               CANTSA = CVD(Mid$(YY$, 89, 8))
               DEPX% = DepOrig%
               LOTEX$ = TRIM$(Mid$(YY$, 107, 12))
               '
               If ESTRAZABLE(CI%) = 2 Then ' SI ES TRAZABLE CON NUMERO DE SERIE
                   ARTICULO$ = CODEXT$(CI%)
                   Call MENSERR(24, "El Artículo: " & ARTICULO$ & " Tiene Nivel de Trazabilidad 2,\No es Posible Realizar la Transferencia por esta Opción")
                   GoTo 1113
               End If
               '
               If LOTEX$ <> "" Then
                  If SaldoLote(CI%, LOTEX$, DEPX%) >= CANTSA Then
                     GoTo 19
                  End If
               End If
               LOTEX$ = LOTESELEC$(CI%, DEPX%, CANTSA)
               HUBOSELOT% = 1
               If LOTEX$ = "" Then
                  GoTo 1113
               End If
19             Call REPLA(YY$, LOTEX$, 69, 12)
               Call REPLA(YY$, LOTEX$, 107, 12)
               Call GRAREG("!BOLREDET", YY$, YX&)
           End If
        Next YX&
        Call CIERRARCH("!BOLREDET")
        If HUBOSELOT% = 1 Then GoTo 1113
        '
        AINDI% = 0
        NUORIT% = (-1)
        TRANAPRO% = 0
        'VALIDACION DE TRANSFERENCIA PENDIENTE DE APROBACION
        If Control$("STOCKS", "TRANSAPRO") = "SI" Then TRANAPRO% = 1
        If TRANAPRO% > 0 Then
            Call COPYSTRU("MOVISTO", "APROTRAN")
            NVM$ = TRIM$(Str$(CMOVDAT_TRAN& + 1)) ' IDENTIFICADOR CMOVDAT& ES UNA FUNCION
            While Len(NVM$) < 6: NVM$ = "0" + NVM$: Wend
            REMI$ = CMOV$ + "." + NVM$
            REMI$ = TRIM$(REMI$)
        End If
        '
        'AGREGUE ESTO POR QUE NO ESTA IMPRIMIENDO EL NUMERO DE TRANSFERENCIA********
        DOPRI$ = "TR." + TRIM$(Str$(NRO&))
        DOSEC$ = "TR." + Mid$(HO$, 8, 1) + Mid$(HO$, 4, 2) + Left$(HO$, 2)
        While Len(DOSEC$) < 10
          DOSEC$ = DOSEC$ + " "
        Wend
                                  '

        For YX& = 1 To ULTREG&("!BOLREDET")
          YY$ = REGLEIDO$("!BOLREDET", YX&)
          CI% = CVI(Mid$(YY$, 83, 2))
          CAN# = CVD(Mid$(YY$, 89, 8))
          LOTE$ = Mid$(YY$, 69, 12)
          VALOUNITA# = 0
          NROPED& = CVD(Mid$(YY$, 163, 8))
          If InStr(EMPREAC$, "ALEJPA") > 0 Then ' SOLO ALEJPA
             VALOUNITA# = PRELISTACT#(CI%, TIMONE%) 'DEVUELVE EL PRECIO DE LISTA ACTIVA Y EL TIPO DE MONEDA DE LA LISTA
          End If
            If TRIM$(LOTE$) = "" Then LOTE$ = Mid$(YY$, 107, 12)
          If CI% > 0 Then
            If CI% <= NUMAS% Then
              If CAN# > 0.005 Then
'                DOSEC$ = "" 'ver si DOSEC$ HAY QUE LIMPIARLO ADENTRO
1120            CANTI = (-1) * CAN#
                NUORIT% = NUORIT% + 1
                If InStr(EMPREAC$, "CAMBACERES") > 0 Then
                  CONDI2$ = "CODIMOVI = 'OF' AND CANTINGRE > 0  AND IDENLOTE = '" & LOTE$ & "'"
                  PESOKG# = XVALO(VALOBADA("MOVISTO", "PESOKG", CONDI2$, "NOMESS"))
                  METROS# = XVALO(VALOBADA("MOVISTO", "METROS", CONDI2$, "NOMESS"))
                End If
                
                If TRANAPRO% > 0 Then 'PENDIEINTE DE APROBACION
                    Call MOVISTOCK2(FEMOVI%, CI%, LOTE$, "TR", REMI$, DOSEC$, "(1) " + REFMO$, 0, "", NC%, DepOrig%, CANTI)
                Else
                    Call MOVISTOK(FEMOVI%, CI%, LOTE$, "TR", DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFMO$, NUORIT%, VALOUNITA#, TIMONE%, 0, DepOrig%, CANTI, NROPED&, 0, "", "", "", 0, TIPOTRA%, , , , METROS#, PESOKG#)
                End If
                
                NUORIT% = NUORIT% + 1
                If TRANAPRO% > 0 Then
                    Call MOVISTOCK2(FEMOVI%, CI%, LOTE$, "TR", REMI$, DOSEC$, "(1) " + REFMO$, 0, "", NC%, DepoDes%, (-1) * CANTI)
                Else
                    Call MOVISTOK(FEMOVI%, CI%, LOTE$, "TR", DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFMO$, NUORIT%, VALOUNITA#, TIMONE%, 0, DepoDes%, (-1) * CANTI, NROPED&, 0, "", "", "", 0, TIPOTRA%)
                End If
              End If
            End If
          End If
        Next YX&
        '
        If TRANAPRO% > 0 Then
            Screen.MousePointer = 1
            Call COMUNI("Transferencia Registrada Pendiente de Aprobación")
            Exit Sub
        End If
        '
        Call CIERRARCH("MOVISTO")
         '\\\OT-5-0447-WAR-08/07/05///
         
        'IMPRESION DEL VALE DE TRANSFERENCIA MULTIPLE
        If FORIPRE$ <> "" Then
           'VALIDACION DE MODO DE IMPRESION
           MODOIMPRE07.Check7.Caption = "Imprime Transferencia"
           MODOIMPRE07.Label1.Caption = "" 'LIMPIA EL LABEL DONDE PONE OK, SI ACEPTA
           MODOIMPRE07.Check8.Value = 0
           Screen.MousePointer = 1
           MODOIMPRE07.Show 1
           CANCELPRINT% = 0
           If MODOIMPRE07.Label1.Caption = "OK" Then
              If MODOIMPRE07.Check8.Value > 0 Then CANCELPRINT% = 1
           End If
           Screen.MousePointer = 11
        '

          CODPARAM$(1) = "NUME-COM": DOCPARAM$(1) = CStr(NRO&)
          CODPARAM$(2) = "FECH-EMI": DOCPARAM$(2) = FECHATEX(FEMOVI%)
          CODPARAM$(3) = "DEPOSITO": DOCPARAM$(3) = DEPOSI$
          CODPARAM$(4) = "DESTINO": DOCPARAM$(4) = DEPODE$
          CODPARAM$(5) = "REF-MAT1": DOCPARAM$(5) = REFMO$
          CAPARDOC% = 5
          '
          CODTABU1$(1) = "COD-MPRI"
          CODTABU1$(2) = "DES-MPRI"
          CODTABU1$(3) = "UNI-MPRI"
          CODTABU1$(4) = "CAN-MPRI"
          CACOLTAB1% = 4
          '
          CAITAB1% = 0
          '
          For YX& = 1 To ULTREG&("!BOLREDET")
            YY$ = REGLEIDO$("!BOLREDET", YX&)
            CI% = CVI(Mid$(YY$, 83, 2))
            CAN# = CVD(Mid$(YY$, 89, 8))
            '
            CAITAB1% = CAITAB1% + 1
              CODMAT$ = TRIM$(CODEXT(CI%))
            TETABU1$(1, CAITAB1%) = CODMAT$
              DESMAT$ = DESCRIT$(CI%)
            TETABU1$(2, CAITAB1%) = DESMAT$
              UNIMAT$ = Unimed$(CI%)
            TETABU1$(3, CAITAB1%) = UNIMAT$
              CANTMAT$ = FORMATNUM(CAN#, "F12.2")
            TETABU1$(4, CAITAB1%) = CANTMAT$
          Next YX&
          '
          Call PRINTDOC("FORTRSTM")
        End If
        '\\\OT-5-0447-WAR-FIN///
        Call COMUNI("Transferencia Procesada")
        '
      End If
      GoTo 100 'MODIFICADO 100 X 110 PARA QUE VUELVA A ACTUALIZAR EL NUMERO DE COMPROBANTE.
    End If
    'CODIGO DE INGRESO POR CODIGO DE BARRA EN ITEM UNICO
'    If InStr(EMPREAC$, "ALEJPA") > 0 Then ' SOLO ALEJPA
'135     OPXBARRAUN% = COMALTER%("Opciones de Carga\\Carga Manual\Código de Barras")
'        If OPXBARRAUN% < 1 Or OPXBARRAUN% > 2 Then GoTo 110
'        '
'        If OPXBARRAUN% = 2 Then
'           LECCODBA.List1.Clear
'137        LECCODBA.Show 1
'           '
'           If XVALO(LECCODBA.List1.ListCount) < 1 Then GoTo 135
'           '
'           VDEF$ = String$(64, 0)
'
'           Call SETULTREG&("!BOLREDET", 0)
'           CI% = LECCODBA.LECT_CODIGO%(LECCODBA.List1, 0)
'           CANTI = LECCODBA.LECT_CANTIDAD(LECCODBA.List1, 0)
'           '
'           FECHA% = HOY(HO$)
'           Call REPLA(VDEF$, MKI$(CI%), 1, 2)
'           Call REPLA(VDEF$, MKI$(CI%), 4, 2)
'           Call REPLA(VDEF$, MKI$(FECHA%), 7, 2)
'           Call REPLA(VDEF$, MKS$(CANTI), 25, 4)
'           '
'           GoTo 140
'        End If
'     End If

140 Screen.MousePointer = 1
150 OpcionItemUnico% = COMALTER%("Opciones:\\Ingreso Manual\Lectura Etiqueta")
    
    VDEF$ = ""
    Call REPLA(VDEF$, Chr(0), 3, 1) ' DEPOSITO SALIDA
    Call REPLA(VDEF$, Chr(0), 6, 1) ' DEPOSITO ENTRADA
    Call REPLA(VDEF$, MKI$(HOY(HO$)), 7, 2) ' FECHA TRANSFERENCIA
    If OpcionItemUnico% = 2 Then
        ' LEER ETIQUETA
        LecturaEtiqueta$ = TRIM$(InputBox("Lectura", "Lectura Etiqueta", ""))
        LecturaEtiqueta$ = REPLACAR$(LecturaEtiqueta$, "'", "-")
        LecturaEtiqueta$ = REPLACAR$(LecturaEtiqueta$, "]", "|")
        If LecturaEtiqueta$ = "" Then GoTo 100
        n% = InStr(1, LecturaEtiqueta$, "|")
        If n% < 1 Then GoTo 100
        CodigoProducto$ = TRIM$(Mid$(LecturaEtiqueta$, 1, n% - 1))
        LoteBobina$ = UCase(TRIM$(Mid$(LecturaEtiqueta$, n% + 1)))
        CodigoProductoInterno% = CODINT%(CodigoProducto$)
        If CodigoProductoInterno% = 0 Then GoTo 100
        
        ' COMPLETA VALORES POR DEFECTO A PARTIR DE LA LECTURA
        Call REPLA(VDEF$, MKI$(CodigoProductoInterno%), 1, 2)
        Call REPLA(VDEF$, MKI$(CodigoProductoInterno%), 4, 2)
        Call REPLA(VDEF$, LoteBobina$, 9, 16)
    End If

    X$ = OBJPLA$("TRANSFESTOCK", VDEF$)
    If X$ = "" Then
       Exit Sub
    End If
    '
    VDEF$ = X$
    TRANSFE$ = X$
    '
    CI1% = CVI(Left$(TRANSFE$, 2))
    DP1% = Asc(Mid$(TRANSFE$, 3, 1))
    CI2% = CVI(Mid$(TRANSFE$, 4, 2))
    DP2% = Asc(Mid$(TRANSFE$, 6, 1))
    FECHA% = CVI(Mid$(TRANSFE$, 7, 2))
      If FECHA% < 1 Then FECHA% = HOY(HO$)
    LOTE$ = TRIM$(Mid$(TRANSFE$, 9, 16))
    CAN = CVS(Mid$(TRANSFE$, 25, 4))
    REFE0$ = TRIM$(Mid$(TRANSFE$, 29, 32))
    '
    'VALIDACION DE FECHA DE TRANSFERENCIA SEGUN CONTROL DEL SETUP DE STOCKS
    If DERACCE%("CAMFETRA/NOMESS", "Validación de Fecha de Transferencia") < 2 Then
        If Control$("STOCKS", "VALFETRA") = "SI" Then
          DIAHO% = HOY(HO$)
          If FECHA% <> DIAHO% Then
             Call COMUNI("No Válido Realizar Transferencia Con Fecha Diferente a la de Hoy")
             GoTo 140
          End If
        End If
    End If

    If ESTRAZABLE(CI1%) = 2 Then '  SE EXCLUYEN LOS MATERIALES CON TRAZABILIDAD = 2
        Call COMUNI("Imposible Transferir desde esta Opción\Materiales con Trazabilidad 2 (Numero de Serie)")
        GoTo 150
    End If
    '
    If CI1% > 0 Then
      If CI1% <= NUMAS% Then
        If DP1% > 0 Then
           If CI2% > 0 Then
              If CI2% <= NUMAS% Then
                If DP2% > 0 Then
                   '\\\OT-08-0842-OD-17/09/08///
                   If CAN >= 0.05 Then
                   '\\\OT-08-0842-OD-FIN///
                      GoTo 155
                   End If
                End If
              End If
           End If
        End If
      End If
    End If
    '

    Call MENSERR(24, "Datos Incompletos !!!\Transferencia Abortada.")
    GoTo 150
    '\\\OT-05-0730-WAR-10/11/05///
    '
155 If ESTRAZABLE(CI1%) = 1 Then
      If LOTE$ = "" Then
         Call MENSERR(24, "Mandatorio Indicar Lote de Origen")
         GoTo 140
      End If
      If SaldoLote(CI1%, LOTE$, DP1%) < CAN - FRACMIN(CI%) / 2 Then
         Call MENSERR(24, "Saldo del Lote Elegido en Deposito de Origen Insuficiente")
         GoTo 140
      End If
    End If
    '
    
160 CAITMO% = 0
    For YX& = 1 To 1
         CI% = CI1%
         CAN = CAN
         If CI% > 0 Then
           If CI% <= NUMAS% Then
             If CAN > 0.005 Then
               CAITMO% = CAITMO% + 1
               CIMOVIM%(CAITMO%) = CI%
               CAMOVIM(CAITMO) = (-1) * CAN
             End If
           End If
         End If
    Next YX&
    CANTMETROSBOBINA# = 0: PESOSALDOBOBINA# = 0
    If InStr(EPAC$, "QUILMES") > 0 Then
        CANTMETROSBOBINA# = SaldoLotexMetros(CI1%, LOTE$, DP1%)
        PESOSALDOBOBINA# = SaldoLotexKilos(CI1%, LOTE$, DP1%)
    End If
    If PUEDEMOVI%(DP1%) < 0 Then GoTo 150 ' CANCELA SALIDA
    '\\\OT-05-0730-WAR-FIN///
    NC% = 0
    VUNI# = 0
    MONE$ = ""
    '
    DOPRI$ = "TR." + TRIM$(Str$(NRO&))
    DOSEC$ = "TR." + Mid$(HO$, 8, 1) + Mid$(HO$, 4, 2) + Left$(HO$, 2)
                              While Len(DOSEC$) < 10
                                DOSEC$ = DOSEC$ + " "
                              Wend
                              '
    COD$ = LTrim$(RTrim$(CODEXT$(CI2%)))
    REFE$ = TRIM$(REFE0$)
    If REFE$ = "" Then
       REFE$ = "TRANSF. A " + COD$ + " - D." + Str$(DP2%)
    End If
    '
    CANTI = (-1) * CAN
    '
    'ANTES DE REALIZAR EL MOVIMIENTO DE TRANSFERENCIA VERIFICA SI ESTA CONFIGURADO
    'QUE LA TRANSFERENCIA SEA CON APROBACION GUARDA EL MOVIMIENTO DE ENTRADA Y DE SALIDA
    'EN EL ARCHIVO DAT TRANSAPRO CON LA ESTRUCTURA DEL MOVISTO
    TRANAPRO% = 0
    If Control$("STOCKS", "TRANSAPRO") = "SI" Then TRANAPRO% = 1
    If TRANAPRO% > 0 Then
'        REFE$ = "De " & DP1% & " A " & DP2%
        Call COPYSTRU("MOVISTO", "APROTRAN")
        NVM$ = TRIM$(Str$(CMOVDAT_TRAN& + 1)) ' IDENTIFICADOR CMOVDAT_TRAN& ES UNA FUNCION
        While Len(NVM$) < 6: NVM$ = "0" + NVM$: Wend
        DOPRI$ = CMOV$ + "." + NVM$
        DOPRI$ = TRIM$(DOPRI$)
        Call MOVISTOCK2(FECHA%, CI1%, LOTE$, CMOV$, DOPRI$, DOSEC$, "(0) " + REFE$, 0, "", NC%, DP1%, CANTI)
    Else
        Call MOVISTOK(FECHA%, CI1%, LOTE$, CMOV$, DOPRILA$, DOPRI$, DOSELA$, DOSEC$, REFE$, 0, VUNI#, MONEI%, NC%, DP1%, CANTI, , , , , , , , , , , CANTMETROSBOBINA#, PESOSALDOBOBINA#)
    End If
    '
    COD$ = LTrim$(RTrim$(CODEXT$(CI1%)))
    REFE$ = TRIM$(REFE0$)
    If REFE$ = "" Then
       REFE$ = "TRANSF. DE " + COD$ + " - D." + Str$(DP1%)
    End If
    '
    'VALIDA EL SEGUNDO MOVIMIENTO DE STOCKS DEPENDIENDO SI REALIZA EL MOVIMIENTO O LO
    'GUARDA PARA PENDIENTE AUTORIACION
    CANTI = CAN
    If TRANAPRO% > 0 Then
       
       Call MOVISTOCK2(FECHA%, CI2%, LOTE$, CMOV$, DOPRI$, DOSEC$, "(0) " + REFE$, 0, "", NC%, DP2%, CANTI)
       Call CIERRARCH("APROTRAN")
       '
       Call COMUNI("Transferencia Registrada Pendiente de Aprobación")
       Exit Sub
    Else
       Call MOVISTOK(FECHA%, CI2%, LOTE$, CMOV$, DOPRILA$, DOPRI$, DOSELA$, DOSEC$, REFE$, 1, VUNI#, MONEI%, NC%, DP2%, CANTI, , , , , , , , , , , CANTMETROSBOBINA, PESOSALDOBOBINA)
       Call CIERRARCH("MOVISTO")
    End If
    '
'IMPRIMIR FORMULARIO

237 If TRIM$(FORIPRE$) <> "" Then
           'VALIDACION DE MODO DE IMPRESION
           MODOIMPRE07.Check7.Caption = "Imprime Transferencia"
           MODOIMPRE07.Label1.Caption = "" 'LIMPIA EL LABEL DONDE PONE OK, SI ACEPTA
           MODOIMPRE07.Check8.Value = 0
           Screen.MousePointer = 1
           MODOIMPRE07.Show 1
           CANCELPRINT% = 0
           If MODOIMPRE07.Label1.Caption = "OK" Then
              If MODOIMPRE07.Check8.Value > 0 Then CANCELPRINT% = 1
           End If
           Screen.MousePointer = 11
             LU$ = LUDAT$
             CODFOR$ = TRIM$(UCase$(FORIPRE$))
             '
             YAC = XVALO(ATRIFORM$(CODFOR$, "DETDESMM"))
             MARIZW = XVALO(ATRIFORM$(CODFOR$, "MARIZQMM"))
             MARSUW = XVALO(ATRIFORM$(CODFOR$, "MARSUPMM"))
             '
             INTERLI = XVALO(ATRIFORM$(CODFOR$, "INTERLIN"))
             If INTERLI < 1 Then
                Call COMUNI("Error de Configuración de Interlinea\en Formulario '" + CODFOR$ + ".FRM'.")
                INTERLI = 1
             End If
             '
             CACOPI% = XVALO(ATRIFORM$(CODFOR$, "CAN-COP"))
             If CACOPI% < 1 Or CACOPI% > 6 Then
                Call COMUNI("Error de Configuración de Cantidad de Copias\en Formulario '" + CODFOR$ + ".FRM'.")
                CACOPI% = 1
             End If
             '
             For COP% = 1 To CACOPI%
               '
               Call CARFORWIN
               '
               FEX = FECHA%
               Call PRINTFORWIN(FECHATEX$(FEX), "FECH-EMI", 0, 0)
               Call PRINTFORWIN(Time$, "HORA-EMI", 0, 0)
               Call PRINTFORWIN(USERNAME$, "USUARIO", 0, 0)
               Call PRINTFORWIN(Str$(NRO&), "NUME-COM", 0, 0)
               Call PRINTFORWIN(TRIM$(CODEXT$(CI1%)), "COD-MAT", 0, YAC)
               Call PRINTFORWIN(TRIM$(CODEXT$(CI2%)), "COD-MAT", 0, YAC + INTERLI)
               Call PRINTFORWIN(DESCRIT$(CI1%), "DES-MAT", 0, YAC)
               Call PRINTFORWIN(DESCRIT$(CI2%), "DES-MAT", 0, YAC + INTERLI)
               Call PRINTFORWIN(Unimed$(CI1%), "UNIMED", 0, YAC)
               Call PRINTFORWIN(Unimed$(CI2%), "UNIMED", 0, YAC + INTERLI)
               Call PRINTFORWIN(ECOARCH$("ECODEP", Chr$(DP1%)), "DEPOSITO", 0, YAC)
               Call PRINTFORWIN(ECOARCH$("ECODEP", Chr$(DP2%)), "DEPOSITO", 0, YAC + INTERLI)
               Call PRINTFORWIN(LOTE$, "NRO-LOTE", 0, YAC)
               CANS$ = CSTRING$(MKS$(CAN), 3, 9, 1, 2)
               Call PRINTFORWIN(CANS$, "CANTIDAD", 0, YAC)
               Call PRINTFORWIN(REFE0$, "REF-MAT1", 0, YAC)
               '
               'Printer.EndDoc
               Call SETSPOOL("ENDDOC", "")
               '
             Next COP%
             Call SAVESPOOL("Transferencia de Stock" + Str$(NRO&), OKX%)
    End If
    '
299 Call COMUNI("Transferencia Completa\y Verificada !!!")
    '
    TRANSFE$ = ""
    GoTo 100
    '
9999
End Sub
Sub TRANSFE_APRO(MODO%)
   '
'   Call COPYSTRU("LILOTSE", "TRANAPRO")
   Call TRALOCAL("TRANAPRO", "")
   FIN& = ULTREG("APROTRAN")
   Call SETULTREG("TRANAPRO", 0)
   Call TRALOCAL("TRANAPRO", "")
   Call FILESORT("APROTRAN", "", 6, 6, "ASC")
   DOPRI_ANT$ = ""
   'CARGO LA LISTA DE SELECCION DESDE EL ARCHIVO APROTRAN QUE ES EL DAT (MOVISTO)
   J& = 0
   TXTX$ = REGBLAN("APROTRAN")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     TXTX$ = REGLEIDO$("APROTRAN", U&)
     
     CIII% = CVI(Mid$(TXTX$, 3, 2))
     If CIII% > 0 Then
        DOPRI$ = TRIM$(Mid$(TXTX$, 23, 10))
       If DOPRI$ <> DOPRI_ANT$ Then
            J& = J& + 1
            DESCRI$ = AJUSTI$(DESCRIT0$(CIII%), 16)
            CANTINGRE = CVD(Mid$(TXTX$, 88, 8))
            CANTSALI = CVD(Mid$(TXTX$, 96, 8))
            REFERE$ = TRIM$(Mid$(TXTX$, 43, 30))
            TIP% = XVALO(Mid$(REFERE$, 2, 1)) ' LEE SI ES 0 O 1 (O UNICO - 1 MULTIPLE)
            NUMUSU% = Asc(Mid$(TXTX$, 126, 1))
            Call REPLA(TTXX$, TRIM$(DOPRI$), 1, 16)
              TEXTOFIJO$ = "Transferencia"
              TIPOTRANSFERENCIA$ = " Item Único "
            If TIP% > 0 Then
              TIPOTRANSFERENCIA$ = " Múltiple "
            End If
            Call REPLA(TTXX$, TEXTOFIJO$, 18, 13)
            Call REPLA(TTXX$, TIPOTRANSFERENCIA$, 31, 19)
            Call REPLA(TTXX$, Mid$(REFERE$, 4), 51, 11) 'MID 4 PARA QUE NO IMPRIMA EL (0) O (1)
            Call REPLA(TTXX$, "", 62, 10)
            If TIP% = 0 Then Call REPLA(TTXX$, FORMATNUM$(CANTSALI, "F10.1"), 62, 10)
            Call REPLA(TTXX$, USERNAM$(NUMUSU%), 74, 16)
            Call REGAPP("!TRANAPRO", TTXX$)
            DOPRI_ANT$ = DOPRI$
       End If
     End If
   Next U&
   Call SETULTREG("!TRANAPRO", J&)
   Call TRACENTRAL("TRANAPRO", "")
   If J& < 1 Then Call COMUNI("No Hay Transferencia Para Autorizar"): Exit Sub
   '
   'PRESENTO LISTA DE SELECCION
50 Screen.MousePointer = 1
   Call SELECHO("TRANAPRO/Cod.Movimiento          Código Artículo             Descripción                                                    Cantidad     Usuario")
   '
   'LEO EL NUMERO DE TRANFERENCIA SELECCIONADO
   DOPRI$ = TRIM$(VALACT1$("TRANAPRO"))
   If DOPRI$ = "" Then Exit Sub
   '
   'MUESTRO LA VENTANA CON EL DETALLE DE LA TRANFERENCIA SELECCIONADA
   Call SETULTREG("!TRANFEAP", 0)
   Y$ = REGBLAN$("!TRANFEAP")
   '
   DepOrig% = 0: DepoDes% = 0
   RGM$ = RECFILT$("APROTRAN", 6, DOPRI$)
   URE& = ULTREG&("APROTRAN")
   For k& = 1 To Len(RGM$) Step 4
      REG& = CVS(Mid$(RGM$, k&, 4))
      TX1$ = REGLEIDO$("APROTRAN", REG&)
      If REG& <= URE& Then
        CIII% = CVI(Mid$(TX1$, 3, 2))
        LOTE$ = Mid$(TX1$, 5, 16)
        REFE$ = TRIM$(Mid$(TX1$, 43, 30))
        CANTINGRE1# = CVD(Mid$(TX1$, 88, 8))
        CANTSAL1# = CVD(Mid$(TX1$, 96, 8))
        CANTTOT# = CANTSAL1# + CANTINGRE1#
        DepOrig% = 0: DepoDes% = 0
        If CANTSAL1# > FRACMIN(CIII%) Then DepOrig% = Asc(Mid$(TX1$, 83, 1)): DEP1% = DepOrig%
        If CANTINGRE1# > FRACMIN(CIII%) Then DepoDes% = Asc(Mid$(TX1$, 83, 1)): DEP2% = DepoDes%
        FF% = CVI(Mid$(TX1$, 1, 2))  'TOMA LA FECHA EN QUE SE GENERO, LA TRANSFERENCIA, NO LA DE LA APROBACION
        REFE$ = TRIM$(Mid$(TX1$, 43, 30))
        Call REPLA(Y$, MKI$(CIII%), 1, 2)
        Call REPLA(Y$, MKD$(0), 3, 8)
        TIP% = XVALO(Mid$(REFE$, 2, 1))
        Call REPLA(Y$, MKD$(CANTTOT#), 3, 8)
        Call REPLA(Y$, LOTE$, 11, 12)
        
        Call REPLA(Y$, Chr$(DepOrig%), 23, 1)
        Call REPLA(Y$, Chr$(DepoDes%), 24, 1)
        Call REGAPP("!TRANFEAP", Y$)
      End If
60 Next k&
   'VISUALIZO LA VENTANA CON EL DETALLE DE LA TRANFERENCIA
   VTB% = VENTABU%("TRANFEAP" + "/NC/TITUPAN=Transferencia de Stock - " & FECHATEX$(FF%))
   '
   'SI ES CONSULTA SE VA
   If MODO% > 0 Then GoTo 50
   
   If VTB% < 0 Or ULTREG&("!TRANFEAP") < 1 Then GoTo 50
   
   'VALIDACION DE USUARIO CASO GABAL
   EPAC$ = UCase$(TRIM$(EMPREAC$))
   If InStr(EPAC$, "GABAL") > 0 Then
        User% = USNBR% Mod 100
        If User% <> 2 And User% <> 13 Then
           Call COMUNI("Usuario No Habilitado Para Autorizar Transferencias")
           Exit Sub
        End If
   Else
      If DERACCE%("APRUETRA", "APROBACION DE TRANFERENCIAS") < 2 Then Exit Sub
   End If
   '
   
    Screen.MousePointer = 11
    '
    '\\\OT-05-0730-WAR-10/11/05///
    CAITMO% = 0: DP99% = 0
    For YX& = 1 To ULTREG&("!TRANFEAP")
         YY$ = REGLEIDO$("!TRANFEAP", YX&)
         CI% = CVI(Mid$(YY$, 1, 2))
         CAN# = CVD(Mid$(YY$, 3, 8))
         DP99% = Asc(Mid$(YY$, 23, 1))
         If DP99% < 1 Then GoTo 75 'SOLO CONSIDERA EL DEPOSITO DE SALIDA
         DP1% = DP99%
         If CI% > 0 Then
           If CI% <= NUMAS% Then
             If CAN# > 0.005 Then
               CAITMO% = CAITMO% + 1
               CIMOVIM%(CAITMO%) = CI%
               CAMOVIM(CAITMO) = (-1) * CAN#
             End If
           End If
         End If
75  Next YX&
    '
    'VALIDA STOCKS
    If PUEDEMOVI%(DP1%) < 0 Then GoTo 50 ' CANCELA SALIDA
    ' VALIDA LOS DATOS INGRESADOS
    '
    For YX& = 1 To ULTREG&("!TRANFEAP")
       YY$ = REGLEIDO$("!TRANFEAP", YX&)
       CI% = CVI(Mid$(YY$, 1, 2))
       If ESTRAZABLE%(CI%) > o Then
           ' OBLIGA A INGRESAR EL LOTE
           ' COMPARA CANTIDAD DESPACHADA CON SALDO DEL LOTE
           CANTSA = CVD(Mid$(YY$, 3, 8))
           DEPX% = Asc(Mid$(YY$, 23, 1))
           If DEPX% < 1 Then GoTo 85
           LOTEX$ = TRIM$(Mid$(YY$, 11, 12))
           '
           If ESTRAZABLE(CI%) = 2 Then ' SI ES TRAZABLE CON NUMERO DE SERIE
               ARTICULO$ = CODEXT$(CI%)
               Call MENSERR(24, "El Artículo: " & ARTICULO$ & " Tiene Nivel de Trazabilidad 2,\No es Posible Realizar la Transferencia por esta Opción")
               GoTo 50
           End If
           '
           If LOTEX$ <> "" Then
              If SaldoLote(CI%, LOTEX$, DEPX%) >= CANTSA Then
                 GoTo 50
              End If
           End If
           LOTEX$ = LOTESELEC$(CI%, DEPX%, CANTSA)
           If LOTEX$ = "" Then
              GoTo 50
           End If
           Call REPLA(YY$, LOTEX$, 11, 12)
           Call GRAREG("!TRANFEAP", YY$, YX&)
85         End If
    Next YX&
    '
    'REALIZA LA TRANFERENCIA
    NUORIT% = 0
    CMOV$ = TRIM$(Control$("", "CODTRAST"))
    If CMOV$ = "" Then CMOV$ = "TR"
    NRO& = ProNroComprobante(CMOV$)
    REFE$ = Mid$(REFE$, 4)
    For YX& = 1 To ULTREG&("!TRANFEAP")
      YY$ = REGLEIDO$("!TRANFEAP", YX&)
      CI% = CVI(Mid$(YY$, 1, 2))
      CAN# = CVD(Mid$(YY$, 3, 8))
      LOTE$ = Mid$(YY$, 11, 12)
      DepOrig% = 0: DepoDes% = 0
      If Asc(Mid$(YY$, 23, 1)) > 0 Then DepOrig% = Asc(Mid$(YY$, 23, 1))
      If Asc(Mid$(YY$, 24, 1)) > 0 Then DepoDes% = Asc(Mid$(YY$, 24, 1))
        VALOUNITA# = 0
        REMI$ = "TR." + TRIM$(Str$(NRO&))
        If CI% > 0 Then
          If CI% <= NUMAS% Then
            If CAN# > 0.005 Then
              DOSEC$ = ""
              NUORIT% = NUORIT% + 1
              If DepOrig% > 0 Then
                Call MOVISTOK(FF%, CI%, LOTE$, CMOV$, REMI$, REMI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VALOUNITA#, TIMONE%, 0, DepOrig%, (-1) * CAN#, 0, 0, "", "", "", 0, TIPOTRA%)
              End If
              If DepoDes% > 0 Then
                 Call MOVISTOK(FF%, CI%, LOTE$, CMOV$, REMI$, REMI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VALOUNITA#, TIMONE%, 0, DepoDes%, CAN#, 0, 0, "", "", "", 0, TIPOTRA%)
              End If
            End If
          End If
        
    End If
    Next YX&

   
''' 'BORRO EL REGISTRO
   Call FILTERDELETE("APROTRAN", 6, AJUSTI$(DOPRI$, 10))
   '
   'GENEARA IMPRESION SEGUN CASO ITEM UNICO O LISTA MULTIPLE
   CANCELPRINT% = 0
   Call LIMPIATETAB
   TIPODOC$ = "Vale Transferencia de Stock"
   NUMEDOC$ = TRIM$(Str$(NRO&))
   '
   If TIP% < 1 Then 'CASO ITEM UNICO
     FORIPRE$ = TRIM$(Control$("", "FORTRAST"))
           If FORIPRE$ <> "" Then
             If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") < 1 Then
               FORIPRE$ = ""
             End If
     End If
        
     If TRIM$(FORIPRE$) <> "" Then
           'VALIDACION DE MODO DE IMPRESION
           MODOIMPRE07.Check7.Caption = "Imprime Transferencia"
           MODOIMPRE07.Label1.Caption = "" 'LIMPIA EL LABEL DONDE PONE OK, SI ACEPTA
           MODOIMPRE07.Check8.Value = 0
           Screen.MousePointer = 1
           MODOIMPRE07.Show 1
           CANCELPRINT% = 0
           If MODOIMPRE07.Label1.Caption = "OK" Then
              If MODOIMPRE07.Check8.Value > 0 Then CANCELPRINT% = 1
           End If
           Screen.MousePointer = 11
             LU$ = LUDAT$
             CODFOR$ = TRIM$(UCase$(FORIPRE$))
             '
             YAC = XVALO(ATRIFORM$(CODFOR$, "DETDESMM"))
             MARIZW = XVALO(ATRIFORM$(CODFOR$, "MARIZQMM"))
             MARSUW = XVALO(ATRIFORM$(CODFOR$, "MARSUPMM"))
             '
             INTERLI = XVALO(ATRIFORM$(CODFOR$, "INTERLIN"))
             If INTERLI < 1 Then
                Call COMUNI("Error de Configuración de Interlinea\en Formulario '" + CODFOR$ + ".FRM'.")
                INTERLI = 1
             End If
             '
             CACOPI% = XVALO(ATRIFORM$(CODFOR$, "CAN-COP"))
             If CACOPI% < 1 Or CACOPI% > 6 Then
                Call COMUNI("Error de Configuración de Cantidad de Copias\en Formulario '" + CODFOR$ + ".FRM'.")
                CACOPI% = 1
             End If
             '
             DepOrig% = 0: DepoDes% = 0
             For YX& = 1 To ULTREG&("!TRANFEAP")
                YY$ = REGLEIDO$("!TRANFEAP", YX&)
                If YX& = 1 Then CI1% = CVI(Mid$(YY$, 1, 2))
                If YX& > 1 Then CI2% = CVI(Mid$(YY$, 1, 2))
                CANS$ = SAFETEXT$(CVD(Mid$(YY$, 3, 8)))
                LOTE$ = Mid$(YY$, 11, 12)
                If Asc(Mid$(YY$, 23, 1)) > 0 Then DepOrig% = Asc(Mid$(YY$, 23, 1))
                If Asc(Mid$(YY$, 24, 1)) > 0 Then DepoDes% = Asc(Mid$(YY$, 24, 1))
             Next YX&
             For COP% = 1 To CACOPI%
               '
               Call CARFORWIN
               '
               FEX = FECHA%
               Call PRINTFORWIN(FECHATEX$(FF%), "FECH-EMI", 0, 0)
               Call PRINTFORWIN(Time$, "HORA-EMI", 0, 0)
               Call PRINTFORWIN(USERNAME$, "USUARIO", 0, 0)
               Call PRINTFORWIN(Str$(NRO&), "NUME-COM", 0, 0)
               Call PRINTFORWIN(TRIM$(CODEXT$(CI1%)), "COD-MAT", 0, YAC)
               Call PRINTFORWIN(TRIM$(CODEXT$(CI2%)), "COD-MAT", 0, YAC + INTERLI)
               Call PRINTFORWIN(DESCRIT$(CI1%), "DES-MAT", 0, YAC)
               Call PRINTFORWIN(DESCRIT$(CI2%), "DES-MAT", 0, YAC + INTERLI)
               Call PRINTFORWIN(Unimed$(CI1%), "UNIMED", 0, YAC)
               Call PRINTFORWIN(Unimed$(CI2%), "UNIMED", 0, YAC + INTERLI)
               Call PRINTFORWIN(ECOARCH$("ECODEP", Chr$(DepOrig%)), "DEPOSITO", 0, YAC)
               Call PRINTFORWIN(ECOARCH$("ECODEP", Chr$(DepoDes%)), "DEPOSITO", 0, YAC + INTERLI)
               Call PRINTFORWIN(LOTE$, "NRO-LOTE", 0, YAC)
               CANS$ = CSTRING$(MKS$(CAN), 3, 9, 1, 2)
               Call PRINTFORWIN(CANS$, "CANTIDAD", 0, YAC)
               Call PRINTFORWIN(REFE$, "REF-MAT1", 0, YAC)
               '
               'Printer.EndDoc
               Call SETSPOOL("ENDDOC", "")
               '
             Next COP%
             Call SAVESPOOL("Transferencia de Stock" + Str$(NRO&), OKX%)
       End If
   End If
   'FIN IMPRESION ITEM UNICO
   '
   'IMPRESION DEL VALE DE TRANSFERENCIA MULTIPLE
   If TIP% > 0 Then 'CASO LISTA MULTIPLE DE TRANFERENCIA
         FORIPRE$ = TRIM$(Control$("", "FORTRSTM"))
         If FORIPRE$ <> "" Then
            LU$ = LUDAT$
            If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
                RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
                Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Vale de Transferencia:\  \'" + RMCC$ + "'.")
                Exit Sub
            End If
         End If
         '
         If FORIPRE$ <> "" Then
           'VALIDACION DE MODO DE IMPRESION
           MODOIMPRE07.Check7.Caption = "Imprime Transferencia"
           MODOIMPRE07.Label1.Caption = "" 'LIMPIA EL LABEL DONDE PONE OK, SI ACEPTA
           MODOIMPRE07.Check8.Value = 0
           Screen.MousePointer = 1
           MODOIMPRE07.Show 1
           CANCELPRINT% = 0
           If MODOIMPRE07.Label1.Caption = "OK" Then
              If MODOIMPRE07.Check8.Value > 0 Then CANCELPRINT% = 1
           End If
           Screen.MousePointer = 11
           '

          CODPARAM$(1) = "NUME-COM": DOCPARAM$(1) = CStr(NRO&)
          CODPARAM$(2) = "FECH-EMI": DOCPARAM$(2) = FECHATEX(FF%)
          CODPARAM$(3) = "DEPOSITO": DOCPARAM$(3) = DEP1%
          CODPARAM$(4) = "DESTINO": DOCPARAM$(4) = DEP2%
          CODPARAM$(5) = "REF-MAT1": DOCPARAM$(5) = REFE$
          CAPARDOC% = 5
          '
          CODTABU1$(1) = "COD-MPRI"
          CODTABU1$(2) = "DES-MPRI"
          CODTABU1$(3) = "UNI-MPRI"
          CODTABU1$(4) = "CAN-MPRI"
          CACOLTAB1% = 4
          '
          CAITAB1% = 0
          '
          For YX& = 1 To ULTREG&("!TRANFEAP")
            YY$ = REGLEIDO$("!TRANFEAP", YX&)
            CI% = CVI(Mid$(YY$, 1, 2))
            CAN# = CVD(Mid$(YY$, 3, 8))
            DepOrig% = Asc(Mid$(YY$, 23, 1))
            '
            If DepOrig% > 0 Then 'PARA QUE SOLO IMPRIMA UNA SOLA LINEA POR CODIGO
                CAITAB1% = CAITAB1% + 1
                  CODMAT$ = TRIM$(CODEXT(CI%))
                TETABU1$(1, CAITAB1%) = CODMAT$
                  DESMAT$ = DESCRIT$(CI%)
                TETABU1$(2, CAITAB1%) = DESMAT$
                  UNIMAT$ = Unimed$(CI%)
                TETABU1$(3, CAITAB1%) = UNIMAT$
                  CANTMAT$ = FORMATNUM(CAN#, "F12.2")
                TETABU1$(4, CAITAB1%) = CANTMAT$
            End If
          Next YX&
          '
          Call PRINTDOC("FORTRSTM")
          '
        End If
   End If
   '
   Screen.MousePointer = 1
   Call COMUNI("Ajuste Número" + Str$(NRO&) + "\Completo y Verificado")
   '
   Call CIERRARCH("*.*")
   CANCELPRINT% = 0
   Call LIMPIATETAB
   
   
End Sub

Sub ANULATRANSFEPEN()
   '
'   Call COPYSTRU("LILOTSE", "TRANAPRO")
   Call SETULTREG("TRANAPRO", 0)
   Call TRALOCAL("TRANAPRO", "")
   Call FILESORT("APROTRAN", "", 6, 6, "ASC")
   DOPRI_ANT$ = ""
   'CARGO LA LISTA DE SELECCION DESDE EL ARCHIVO APROTRAN QUE ES EL DAT (MOVISTO)
10 Screen.MousePointer = 11
   TXTX$ = REGBLAN("APROTRAN")
   FIN& = ULTREG("APROTRAN")
   J& = 0
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     TXTX$ = REGLEIDO$("APROTRAN", U&)
     CIII% = CVI(Mid$(TXTX$, 3, 2))
     If CIII% > 0 Then
        DOPRI$ = TRIM$(Mid$(TXTX$, 23, 10))
        If DOPRI$ <> DOPRI_ANT$ Then
            J& = J& + 1
            DESCRI$ = AJUSTI$(DESCRIT0$(CIII%), 16)
            CANTINGRE = CVD(Mid$(TXTX$, 88, 8))
            CANTSALI = CVD(Mid$(TXTX$, 96, 8))
            REFERE$ = TRIM$(Mid$(TXTX$, 43, 30))
            TIP% = XVALO(Mid$(REFERE$, 2, 1)) ' LEE SI ES 0 O 1 (O UNICO - 1 MULTIPLE)
            NUMUSU% = Asc(Mid$(TXTX$, 126, 1))
            Call REPLA(TTXX$, TRIM$(DOPRI$), 1, 16)
              TEXTOFIJO$ = "Transferencia"
              TIPOTRANSFERENCIA$ = " Item Único "
            If TIP% > 0 Then
              TIPOTRANSFERENCIA$ = " Múltiple "
            End If
            Call REPLA(TTXX$, TEXTOFIJO$, 18, 13)
            Call REPLA(TTXX$, TIPOTRANSFERENCIA$, 31, 19)
            Call REPLA(TTXX$, Mid$(REFERE$, 4), 51, 11) 'MID 4 PARA QUE NO IMPRIMA EL (0) O (1)
            Call REPLA(TTXX$, "", 62, 10)
            If TIP% = 0 Then Call REPLA(TTXX$, FORMATNUM$(CANTSALI, "F10.1"), 62, 10)
            Call REPLA(TTXX$, USERNAM$(NUMUSU%), 74, 16)
            Call REGAPP("!TRANAPRO", TTXX$)
            DOPRI_ANT$ = DOPRI$
       End If
     End If
   Next U&
   Call SETULTREG("!TRANAPRO", J&)
   Call TRACENTRAL("TRANAPRO", "")
   If J& < 1 Then Call COMUNI("No Hay Transferencia Para Anular"): Exit Sub
   '
   'PRESENTO LISTA DE SELECCION
   Screen.MousePointer = 1
50 Call SELECHO("TRANAPRO/Cod.Movimiento          Código Artículo             Descripción                                                    Cantidad     Usuario")
   '
   'LEO EL NUMERO DE TRANFERENCIA SELECCIONADO
   DOPRI$ = TRIM$(VALACT1$("TRANAPRO"))
   If DOPRI$ = "" Then Exit Sub
   '
   OPX% = COMALTER%("¿Esta Seguro que Desea Anular la Transferencia '" & DOPRI$ & "'?\\No, Cancelar\Si, Anular")
   If OPX% = 2 Then
        EPAC$ = UCase$(TRIM$(EMPREAC$))
        If InStr(EPAC$, "GABAL") > 0 Then
             User% = USNBR% Mod 100
             If User% <> 2 And User% <> 13 Then
                Call COMUNI("Usuario No Habilitado Para Anular Transferencias")
                Exit Sub
             End If
        Else
           If DERACCE%("APRUETRA", "APROBACION DE TRANFERENCIAS") < 2 Then Exit Sub
        End If
   Else
        GoTo 50
   End If
   '
   DOPRI$ = AJUSTI$(DOPRI$, 10)
   Call FILTERDELETE("APROTRAN", 6, DOPRI$)
   Call FRESHECHO("APROTRAN", "")
   Call COMUNI("Transferencia Anulada")
   '
End Sub
'
Sub MOVISTOCK2(FECHA%, CI%, LOTE$, CMOV$, DOPRI$, DOSEC$, REFE$, VUNI#, MONE$, NC%, DEPO%, CANTI)
   '
   Static CTX%, REBLAN$, RELIBR$, UREMOVI&, DECONSUM$
   '
''   If CANTI < 0 Then
''     If DECONSUM$ = "" Then
''       DECONSUM$ = "MANUA"
''       If CONTROL$("", "DECONSUM") = "AUTOM" Then
''         DECONSUM$ = "AUTOM"
''       End If
''     End If
''   End If
   '
   SMP = Screen.MousePointer
   UREMAS& = ULTREG&("MASTER")
   '
'''   If ULTREG&("APROTRAN") < UREMOVI& Then
'''      Call CIERRARCH("*.*")
'''      Call MENSERR(24, "Imposible Registrar este Movimiento.\  \Cierre la Aplicación y Vuelva a Intentarlo.")
'''      Call FINAL("")
'''   End If
   '
   User% = USNBR% Mod 100
   If User% < 1 Then User% = 1
   Call CARDEPOS
   '
'''   If CTX% < 1 Then
'''      REBLAN$ = REGBLAN$("APROTRAN")
'''      RELIBR$ = String$(128, 45)
'''      Call COPYSTRU("APROTRAN", "BKMOVIST")
'''      CTX% = 1
'''   End If
   '
   ' ********** ARMAR REGISTRO DE MOVIMIENTO **********
9  DPRX$ = DOPRI$
10 PPXX% = InStr(DPRX$, "-0")
   If PPXX% > 0 Then
      DPRX$ = Left$(DPRX$, PPXX%) + Mid$(DPRX$, PPXX% + 2)
      GoTo 10
   End If
   DOPRIX$ = AJUSTI$(DPRX$, 10)
   '
   DPRX$ = DOSEC$
20 PPXX% = InStr(DPRX$, "-0")
   If PPXX% > 0 Then
      DPRX$ = Left$(DPRX$, PPXX%) + Mid$(DPRX$, PPXX% + 2)
      GoTo 20
   End If
   DOSECX$ = AJUSTI$(DPRX$, 10)
   '
   NCLIE% = NC%: NPROV% = 0
   If NC% < 0 Then
      NCLIE% = 0: NPROV% = Abs(NC%)
   End If
   '
   Entrada# = CDbl(CANTI): SALIDA# = 0
   If CANTI < 0 Then
      Entrada# = 0: SALIDA# = CDbl(Abs(CANTI))
   End If
   CANACU = SALDETOT + CANTI
   '
   depomovi% = DEPO%
'''   If DECONSUM$ = "AUTOM" Then
'''     If CANTI < 0 Then
'''       DEPII% = DEPO%
'''       DEPOMOVI% = DEPOSIBLE%(CI%, CANTI, DEPII%)
'''     End If
'''   End If
   '
   RECORD$ = REBLAN$
   Call REPLA(RECORD$, MKI$(FECHA%), 1, 2)
   Call REPLA(RECORD$, MKI$(CI%), 3, 2)
   Call REPLA(RECORD$, LOTE$, 5, 16)
   Call REPLA(RECORD$, CMOV$, 21, 2)
   If CMOV$ = "*R" Then
        Call REPLA(RECORD$, "RP", 21, 2)      ' reparación
      ElseIf CMOV$ = "*D" Then
        Call REPLA(RECORD$, "DR", 21, 2)      ' DESPACHO reparación
   End If
   Call REPLA(RECORD$, DOPRIX$, 23, 10)
   Call REPLA(RECORD$, DOSECX$, 33, 10)
   Call REPLA(RECORD$, REFE$, 43, 30)
   Call REPLA(RECORD$, MKD$(VUNI#), 73, 8)
   Call REPLA(RECORD$, MONE$, 81, 2)
   Call REPLA(RECORD$, Chr$(depomovi%), 83, 1)
   Call REPLA(RECORD$, MKI$(NCLIE%), 84, 2)
   Call REPLA(RECORD$, MKI$(NPROV%), 86, 2)
   Call REPLA(RECORD$, MKD$(Entrada#), 88, 8)
   Call REPLA(RECORD$, MKD$(SALIDA#), 96, 8)
   Call REPLA(RECORD$, String$(8, 0), 104, 8)
   Call REPLA(RECORD$, String$(6, 0), 112, 6)
   Call REPLA(RECORD$, MKS$(CANACU), 118, 4)
   '
'''   If UCase$(CMOV$) = "PP" Or CMOV$ = "DM" Then
'''     CAVERI$ = DOSECX$
'''     CAVERJ$ = MKS$(Val(Left$(CAVERI$, 6))) + MKI$(Val(Mid$(CAVERI$, 8, 3)))
'''     Call REPLA(RECORD$, CAVERJ$, 112, 6)
'''   End If
   '
   HO% = HOY(HOS$)
   HP = 100 * Val(Left$(Time$, 2)) + Val(Mid$(Time$, 4, 2))
   '
   Call REPLA(RECORD$, MKI$(HO%), 122, 2)
   Call REPLA(RECORD$, MKI$(Int(HP)), 124, 2)
   Call REPLA(RECORD$, Chr$(User%), 126, 1)
   Call REPLA(RECORD$, String$(2, 0), 127, 2)
   '
   Dim BUGRA As String * 140
'''   If Left$(CMOV$, 1) <> "*" Then
   Call REGAPP("APROTRAN", RECORD$)
   REMOVIST& = ULTREG&("APROTRAN")
   If REMOVIST& > UREMOVI& Then UREMOVI& = REMOVIST&
   '
   NKX% = FreeFile
   Open LUCOM$ + "\APROTRAN.LOC" For Random Access Read Write As #NKX% Len = 140
   RECONO& = 1 + LOF(NKX%) / 140
   BUGRA$ = "APROTRAN " + MKS$(REMOVIST&) + RECORD$
   Put #NKX%, RECONO&, BUGRA$
   Close #NKX%
   '
   If CI% > 0 Then FLAGMOVI%(CI%) = 1
         '
'''     Else               ' reparaciones
'''         Call COPYSTRU("APROTRAN", "MOVIREPA")
'''         Call REGAPP("MOVIREPA", RECORD$)
'''         REMOVIRE& = ULTREG&("MOVIREPA")
'''         Call CIERRARCH("MOVIREPA")
'''         '
'''         NKX% = FreeFile
'''         Open LUCOM$ + "\APROTRAN.LOC" For Random Access Read Write As #NKX% Len = 140
'''         RECONO& = 1 + LOF(NKX%) / 140
'''         BUGRA$ = "MOVIREPA" + MKS$(REMOVIRE&) + RECORD$
'''         Put #NKX%, RECONO&, BUGRA$
'''         Close #NKX%
'''         '
'''   End If
   '
   Screen.MousePointer = SMP
   '
End Sub



