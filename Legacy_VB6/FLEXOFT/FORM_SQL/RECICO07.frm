VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form RECICO07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recibo de Cobranza - Opciones"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4815
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      ItemData        =   "RECICO07.frx":0000
      Left            =   120
      List            =   "RECICO07.frx":0002
      TabIndex        =   30
      Top             =   7320
      Width           =   10410
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
      Height          =   1185
      ItemData        =   "RECICO07.frx":0004
      Left            =   120
      List            =   "RECICO07.frx":0006
      TabIndex        =   29
      Top             =   5760
      Width           =   10410
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2520
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   26
      Top             =   420
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00624E28&
      ForeColor       =   &H80000008&
      Height          =   5580
      Left            =   3960
      ScaleHeight     =   5550
      ScaleWidth      =   1350
      TabIndex        =   16
      Top             =   -210
      Width           =   1380
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   105
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   31
         Top             =   4500
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -120
            TabIndex        =   32
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command20 
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
         Picture         =   "RECICO07.frx":0008
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   980
         Width           =   650
      End
      Begin VB.CommandButton AGenRep 
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
         Picture         =   "RECICO07.frx":0312
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
         Top             =   2350
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command6 
         Caption         =   "D-C"
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
         Picture         =   "RECICO07.frx":061C
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Listado Aplicación Diferencias de Cambio"
         Top             =   3740
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command1 
         Height          =   690
         Left            =   105
         Picture         =   "RECICO07.frx":0926
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Re-Imputación Contable"
         Top             =   3050
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command2 
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
         Picture         =   "RECICO07.frx":0A70
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   300
         Width           =   650
      End
      Begin VB.CommandButton Command8 
         Caption         =   "List"
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
         Picture         =   "RECICO07.frx":0BBA
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Listados de Recibos Registrados"
         Top             =   1670
         Visible         =   0   'False
         Width           =   650
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2D3C0&
      Caption         =   "OPERACIONES ESPECIALES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4560
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   20
      Width           =   3735
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   1020
         ScaleHeight     =   615
         ScaleWidth      =   2400
         TabIndex        =   20
         Top             =   2520
         Width           =   2430
         Begin VB.CommandButton Command10 
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
            Left            =   0
            Picture         =   "RECICO07.frx":0EC4
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Compensación Cliente-Proveedor"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Compensación Cliente-Proveedor"
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
            Left            =   840
            TabIndex        =   22
            Top             =   80
            Width           =   1530
         End
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   480
         ScaleHeight     =   615
         ScaleWidth      =   2400
         TabIndex        =   13
         Top             =   3600
         Width           =   2430
         Begin VB.CommandButton Command5 
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
            Left            =   0
            Picture         =   "RECICO07.frx":11CE
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Ajuste Directo de Saldos Deudores"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Saldos Iniciales en Cuentas Corrientes"
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
            Left            =   700
            TabIndex        =   15
            Top             =   75
            Width           =   1635
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   810
         ScaleHeight     =   615
         ScaleWidth      =   2400
         TabIndex        =   10
         Top             =   1995
         Width           =   2430
         Begin VB.CommandButton Command7 
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
            Left            =   1920
            Picture         =   "RECICO07.frx":1318
            Style           =   1  'Graphical
            TabIndex        =   33
            ToolTipText     =   "Anular Ajustes "
            Top             =   0
            Width           =   560
         End
         Begin VB.CommandButton Command4 
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
            Left            =   0
            Picture         =   "RECICO07.frx":1462
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Ajuste Directo de Saldos Deudores"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Ajuste de Saldos Deudores"
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
            Left            =   600
            TabIndex        =   12
            Top             =   75
            Width           =   1530
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   600
         ScaleHeight     =   615
         ScaleWidth      =   2400
         TabIndex        =   7
         Top             =   1470
         Width           =   2430
         Begin VB.CommandButton Command11 
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
            Left            =   1850
            Picture         =   "RECICO07.frx":15AC
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Visualizar Imputación de Créditos y Anticipos"
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
            Height          =   630
            Left            =   0
            Picture         =   "RECICO07.frx":18B6
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Imputación y Des-Aplicación de Créditos y Anticipos"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Imputación de Creditos"
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
            Left            =   800
            TabIndex        =   9
            Top             =   80
            Width           =   1110
         End
      End
      Begin VB.PictureBox Picture8 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   390
         ScaleHeight     =   615
         ScaleWidth      =   2400
         TabIndex        =   1
         Top             =   945
         Width           =   2430
         Begin VB.CommandButton Command23 
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
            Left            =   0
            Picture         =   "RECICO07.frx":1A00
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Imprimir Copia de Recibo ya Emitido"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Ver por Pantalla y Re-Imprimir"
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
            Left            =   840
            TabIndex        =   3
            Top             =   80
            Width           =   1500
         End
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   180
         ScaleHeight     =   615
         ScaleWidth      =   2400
         TabIndex        =   4
         Top             =   420
         Width           =   2430
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
            Left            =   0
            Picture         =   "RECICO07.frx":1D0A
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Anulación de Recibo Emitido"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Anulación de Recibo Cobranza"
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
            Left            =   840
            TabIndex        =   6
            Top             =   80
            Width           =   1530
         End
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   4095
         Y1              =   3380
         Y2              =   3380
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "RECICO07.frx":2014
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   60
      OleObjectBlob   =   "RECICO07.frx":2248
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "RECICO07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MODOMONEI%

Function ENSADEUPEN#(DOCU$, NC)
    SQLX$ = "SELECT SUM(SALDACRE) AS IMPOAP FROM SADEUPEN WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE NUCLIEN = " & NC
    SQLX$ = SQLX$ + "  AND CODICOM_LAR = '" & DOCU$ & "'"
    SQLX$ = SQLX$ + "  GROUP BY CODICOM_LAR"
    Set RS = READSET(SQLX$)
    ENSADEUPEN# = XVALO(RS!IMPOAP)
    RS.Close
    Set RS = Nothing
End Function
Function FAC_ENSADEUPEN#(DOCU$, NC)
    SQLX$ = "SELECT SUM(SALDDEBE) AS IMPOAP FROM SADEUPEN WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE NUCLIEN = " & NC
    SQLX$ = SQLX$ + "  AND CODICOM_LAR = '" & DOCU$ & "'"
    SQLX$ = SQLX$ + "  GROUP BY CODICOM_LAR"
    Set RS = READSET(SQLX$)
    FAC_ENSADEUPEN# = XVALO(RS!IMPOAP)
    RS.Close
    Set RS = Nothing
End Function

Function APLICADOENAPLICOBRA#(DOCU$, NC)
    SQLX$ = "SELECT SUM(IMPOAPLI) AS IMPOAP FROM APLICOBRA WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE NUCLIEN = " & NC
    SQLX$ = SQLX$ + "  AND CoDocAP_Lar = '" & DOCU$ & "'"
    SQLX$ = SQLX$ + "  GROUP BY CoDocAP_Lar"
    Set RS = READSET(SQLX$)
    APLICADOENAPLICOBRA# = XVALO(RS!IMPOAP)
    RS.Close
    Set RS = Nothing
End Function
Function FACT_APLICADOENAPLICOBRA#(DOCU$, NC)
    SQLX$ = "SELECT SUM(IMPOAPLI) AS IMPOAP FROM APLICOBRA WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE NUCLIEN = " & NC
    SQLX$ = SQLX$ + "  AND CoDocOr_Lar = '" & DOCU$ & "'"
    SQLX$ = SQLX$ + "  GROUP BY CoDocOr_Lar"
    Set RS = READSET(SQLX$)
    FACT_APLICADOENAPLICOBRA# = XVALO(RS!IMPOAP)
    RS.Close
    Set RS = Nothing
End Function

Function EXISTE_ENAPLICOBRA%(NCLI, RECIBOX$, FEMI%, IMPORIGRECIBO#)
   SQLX$ = "SELECT  SUM(IMPOAPLI) AS IMPORTEAPLICADO FROM APLICOBRA WITH(NOLOCK) "
   SQLX$ = SQLX$ + " WHERE NUCLIEN = " & NCLI
   SQLX$ = SQLX$ + " AND CODOCAP_LAR = '" & RECIBOX$ & "'"
   Set RS = READSET(SQLX$)
   With RS
     Do While Not .EOF
       
       Aplicado# = XVALO(!ImporteAplicado)
       If IMPORIGRECIBO# > Aplicado# + 0.05 Then
          Condi$ = "NUCLIEN = " & NCLI & " AND CODICOM_LAR = '" & RECIBOX$ & "'"
          EXISTEX% = CantRegistros&("SADEUPEN", Condi$, "NOMESS")
          If EXISTEX% > 0 Then
             SALACRE_SADEUPEN# = XVALO(VALOBADA("SADEUPEN", "SUM(SALDACRE)", Condi$, "NOMESS"))
''             If IMPORIGRECIBO# + 0.05 > SALACRE_SADEUPEN# Then
'' VER CUANDO REALIZO EL TEMA DE LAS FACTURAS SI NO ME CONVIENE BORRAR LAS FACTURAS PENDIENTES POR QUE DESPUES LAS VOY A TERMINAR AGREGANDO EN LA RUTINA SIGUIENTE.
''               'ACTUALIZO A SADEUPEN
''             End If
          Else
             '
            NumeClie = NCLI
            CoComCor$ = Left$(RECIBOX$, 3) & Mid$(RECIBOX$, 13, 18) & "-00"
            CoComLar$ = RECIBOX$
            IBruTot# = (-1) * (IMPORIGRECIBO#)
            IIvaTot# = 0
            NuorVen% = 0
            FechEmi = CVDAT(FEMI%)
            FeVenc = CVDAT(FEMI%)
            ImVenc# = (-1) * IMPORIGRECIBO#
            Call GRADEUPEN
    End If

          End If
       
       .MoveNext
     Loop
   End With
   
   
End Function
Sub GRABAFACTURAENSADEUPEN(NCLI, RECIBOX$, FEMI%, IMPOORIGFACTURA#, CONDIPAG)
   SQLX$ = "SELECT  SUM(IMPOAPLI) AS IMPORTEAPLICADO FROM APLICOBRA WITH(NOLOCK) "
   SQLX$ = SQLX$ + " WHERE NUCLIEN = " & NCLI
   SQLX$ = SQLX$ + " AND CODOCOR_LAR = '" & RECIBOX$ & "'"
   Set RS = READSET(SQLX$)
   With RS
     If .EOF Then
          Condi$ = "NUCLIEN = " & NCLI & " AND CODICOM_LAR = '" & RECIBOX$ & "'"
          EXISTEX% = CantRegistros&("SADEUPEN", Condi$, "NOMESS")
          If EXISTEX% < 1 Then
             TIPOX% = 2
             If Left$(RECIBOX$, 2) = "ND" Then TIPOX% = 3
             If Left$(RECIBOX$, 2) = "NC" Then TIPOX% = 6
             letx$ = Right(RECIBOX$, 1)
             VARI% = 1
             If letx$ = "B" Then VARI% = 2
             If letx$ = "E" Then VARI% = 3
             NROX1$ = TRIM$(Mid$(RECIBOX$, 11, 6))
            FEMIX1% = FEMI%
            CoComCor$ = Left$(NROX$, 2) & "-" & NROX1$ & "-" & letx$ 'SIN PREFIJO
            CoComLar$ = RECIBOX$
            NumeClie = NCLI
            NumeProv% = 0
            FechEmi = CVDAT(FEMI%)
            Referen$ = rasocli$(NCLI)
            NumeCom = XVALO(NROX1$)
            Observa$ = rasocli$(NCLI)
            '
            IBruTot# = IMPOORIGFACTURA#
            If CONDIPAX% < 1 Then
              CONPAX% = CPAGCLI%(NCLI)
            End If
            Call FORFASQL.GRADEUFA(TIPOX%, VARI%, NROX1$, NCLI, FEMIX1%, CONPAX%, IMPOORIGFACTURA#)
          End If
     End If
   End With

'     Do While Not .EOF
       'SI LA FACTURA NO EXISTE EN APLICOBRA DEBE ESTAR EN SADEUPEN SI NO ESTA HAY QUE AGREGARLA.
       'SI LA FACTURA EXISTE EN APLICOBRA SUMO EL  IMPOAPLI SI EL IMPORTE DE LA FACTURA ES >  A LA SUMA DE IMPOAPLI
       'ME FIJO SI EXISTE EN SADEUPEN Y SI EXISTE SIGO DE LARGO
       'SI NO EXISTE EN SADEUPEN HAGO UN GRADEUFA.
       
'       APLICADO# = XVALO(!IMPORTEAPLICADO)
'       If IMPOORIGFACTURA# > APLICADO# + 0.05 Then
'          CONDI$ = "NUCLIEN = " & NCLI & " AND CODICOM_LAR = '" & RECIBOX$ & "'"
'          EXISTEX% = CantRegistros&("SADEUPEN", CONDI$, "NOMESS")
'          If EXISTEX% > 0 Then
'             SALACRE_SADEUPEN# = XVALO(VALOBADA("SADEUPEN", "SUM(SALDDEBE)", CONDI$, "NOMESS"))
'''             If IMPOORIGFACTURA# + 0.05 > SALACRE_SADEUPEN# Then
''' VER CUANDO REALIZO EL TEMA DE LAS FACTURAS SI NO ME CONVIENE BORRAR LAS FACTURAS PENDIENTES POR QUE DESPUES LAS VOY A TERMINAR AGREGANDO EN LA RUTINA SIGUIENTE.
'''               'ACTUALIZO A SADEUPEN
'''             End If
'          Else
'             '
'            NumeClie = NCLI
'            CoComCor$ = Left$(RECIBOX$, 3) & Mid$(RECIBOX$, 13, 18) & "-00"
'            CoComLar$ = RECIBOX$
'            IBruTot# = (-1) * (IMPOORIGFACTURA#)
'            IIvaTot# = 0
'            NuorVen% = 0
'            FechEmi = CVDAT(FEMI%)
'            FeVenc = CVDAT(FEMI%)
'            ImVenc# = (-1) * IMPOORIGFACTURA#
'            Call GRADEUPEN
'    End If
'
'          End If
'
'       .MoveNext
'     Loop
'   End With
   
   
End Sub


Sub AGenRep_Click()
    '
    AGenRep.Enabled = False
    Call ACONSUL_PRE(TRIM$(UCase$(Name)))
    AGenRep.Enabled = True
    '
End Sub

Private Sub AGenRep_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   '
   If Button = 2 Then
      Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))
   End If
   '
End Sub

'\\\OT-06-0409-MH-26/09/06///
Sub Command1_Click()
    '
    Command1.Enabled = False
    'jandy SQL
    'JANDY SE COPIO LA RUTINA DE COMMAND8CLICK(LIST)
    'PARA Q FUNCIONE EL BOTON CANCEL
5   OPXX% = ALTERNA%("Reimputación de Comprobantes\Control de Imputaciones")
    If OPXX% < 1 Or OPXX% > 2 Then GoTo 99
    If OPXX% = 2 Then 'CONTROL DE IMPUTACIONES
      Call CONTROLAPLI: GoTo 99
    End If ' SI OPXX% NO ES 2 Y LLEGO ACA ES 1, ENTONCES DIRECTAMENTE EJECUTA LA RE-IMPUTACION
    '
    If DERACCE%("REIMPUVE", "Re-Imputación de Comprobantes de Ventas") < 2 Then
      GoTo 99
    End If
   
8   Call DESHASFECHA(DES%, HAS%, PERIO$) ' RE- IMPUTACION
    If PERIO$ = "" Then
      GoTo 5
    End If

    DESDEX = FETEXCOR1$(DES%)
    HASTAX = FETEXCOR1$(HAS%)
    
    '
10  Call SELECHO("DEUDOR1")
    NCLI = XVALO(VALACT1$("DEUDOR1"))
    If NCLI < 1 Then GoTo 8
    '
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE (TipoMovim = 'RCOB')"
    'OR TIPOMOVIM = 'AJUD'
    SQLX$ = SQLX$ + "AND NuCLIEN = " & NCLI & " "
    SQLX$ = SQLX$ + "AND FechConta >= '" & DESDEX & "' "
    SQLX$ = SQLX$ + "AND FechConta <= '" & HASTAX & "' "
    SQLX$ = SQLX$ + "Order by CODICOM_LAR ASC "
    '
    'Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim CABATEMP As ADODB.Recordset
    Set CABATEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    JJ& = 0
    With CABATEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        Do While Not .EOF
          IMPO# = !ValAbsComp
          If IMPO# >= 0.005 Then
            IMPOS$ = NUMSTRI$(IMPO#, 14, 2, 1)
            NPX& = XVALO(Mid$(!codicom_lar, 11, 8))
            Y$ = MKS$(JJ& + 1) + AJUSTI$(!codicom_lar, 18) + IMPOS$
            JJ& = JJ& + 1
            Call GRAREG("!SELCUERE", Y$, JJ&)
          End If
        .MoveNext
        Loop
      End If
    End With
    CABATEMP.Close
    Set CABATEMP = Nothing
    
    Call SETULTREG("!SELCUERE", JJ&)
    Screen.MousePointer = 1
    '
    If ULTREG&("!SELCUERE") < 1 Then
       Call MENSERR(24, "La Cuenta Elegida no Tiene\Comprobantes de Venta\que Puedan Re-Imputarse.")
       GoTo 10
    End If
    '
    Call CIERRARCH("!SELCUERE")
    Call FRESHECHO("!SELCUERE")
    '
16  Call SELECHO("!SELCUERE")
    If VALACT1$("!SELCUERE") = "" Then GoTo 10
    REGCUEC& = XVALO(VALACT1$("!SELCUERE"))
    If REGCUEC& < 1 Then GoTo 10
    '
    DOCULA$ = TRIM$(Left$(VALACT2$("!SELCUERE"), 18))
    IDOCUM# = XVALO(Mid$(VALACT2$("!SELCUERE"), 19, 14))
    '
20  ATRI$ = DOCULA$ + " - $ " + TRIM$(Mid$(VALACT2$("!SELCUERE"), 19, 14)) + " - " + TRIM$(rasocli$(NCLI))
    '
    Call BLANARCH("!REIMPUTA")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE FechConta >= '" & DESDEX & "' "
    SQLX$ = SQLX$ + "AND FechConta <= '" & HASTAX & "' "
    SQLX$ = SQLX$ + "AND CodiCom_Lar = '" & DOCULA$ & "' "
    SQLX$ = SQLX$ + "AND NuClien = " & NCLI & " "
    'SQLX$ = SQLX$ + "AND Tipomovim<>'RCOB'" ' SE MODIFICA PARA INCLUIR LOS REDONDEOS QUE DISMINUYEN
    'sqlx$ = sqlx$ + "AND IDebeComp > 0 "   ' Y QUE ESTAN DEL LADO DEL DEBE
    '
    'Set CABASEL = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
    Dim CABASEL As ADODB.Recordset
    Set CABASEL = New ADODB.Recordset
25  On Error Resume Next
    CABASEL.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
    End If
    On Error GoTo 0
    SUIMPU# = 0
    '
    With CABASEL
      On Error Resume Next
      .MoveFirst
      If Err Then
        On Error GoTo 0
        Call MENSERR(24, "Imposible Editar Imputación Contable\del Comprobante Elegido")
        GoTo 10
      End If
      '
      Do While Not .EOF
        CUE% = !CUECONTAI
        IPA# = !idebecomp - !ihabecomp
        Y$ = REGBLAN$("!REIMPUTA")
        Call REPLA(Y$, CODCUE$(CUE%), 1, 12)
        Call REPLA(Y$, MKD$(IPA#), 13, 8)
        Call REPLA(Y$, MKI$(CUE%), 29, 2)
        Call REGAPP("!REIMPUTA", Y$)
        SUIMPU# = SUIMPU# + IPA#
        .MoveNext
      Loop
    End With
    
    Call CIERRARCH("!REIMPUTA")
    '
    'If Abs(SUIMPU# - IDOCUM#) >= 0.005 Then
    If Abs(SUIMPU#) >= 0.005 Then
      Call MENSERR(24, "Imposible Editar Imputación Contable.\Comprobante No Balanceado.\  \Debe Anularse y Registrarse Nuevamente.")
      GoTo 10
    End If
    '
30  ATRI$ = "TITUPAN=Re-Imputación Contable: " + ATRI$
    VTB% = VENTABU%("REIMPUTA/" + ATRI$)
    If VTB% < 0 Then GoTo 10
    '
    SUMANE# = 0
    For U& = 1 To ULTREG&("!REIMPUTA")
      Y$ = REGLEIDO$("!REIMPUTA", U&)
      CUENUE% = CUEINT%(Left$(Y$, 12))
      IMPOX# = CVD(Mid$(Y$, 13, 8))
      SUMANE# = SUMANE# + IMPOX#
      '
      If CUENUE% < 1 Then
        Call MENSERR(24, "Numero de Cuenta no Valido\en Registro Numero " + TRIM$(Str$(U&)))
        GoTo 30
      End If
    Next U&
    'If Abs(SUMANE# - IDOCUM#) >= 0.005 Then
    If Abs(SUMANE#) >= 0.005 Then
      Call MENSERR(24, "Comprobante no Balanceado")
      GoTo 20
    End If
    '
    REIMPU% = 0
    For U& = 1 To ULTREG&("!REIMPUTA")
      Y$ = REGLEIDO$("!REIMPUTA", U&)
      CUENUE% = CUEINT%(Left$(Y$, 12))
      IMPOX# = CVD(Mid$(Y$, 13, 8))
      CUEVIE% = CVI(Mid$(Y$, 29, 2))
      '
      'JANDY VER PERFORMANCE SE ABRIO UN RECORDSET PESIMISTA
      'JANDY VER DE PASAR A UPDATE...FLEXCONN.EXECUTE
      
      With CABASEL
        .MoveFirst
        Do While Not .EOF
32        CUE% = !CUECONTAI
          If Abs(!idebecomp - !ihabecomp - IMPOX#) < 0.005 Then
            If CUE% = CUEVIE% Then
             ' .Edit
              !CUECONTAI = CUENUE%
                If CUENUE% <> CUEICARCH% Then
                  !Status = 0
                End If
              .Update
              REIMPU% = 1
              GoTo 35
            End If
          End If
        .MoveNext
        Loop
      End With
      '
35  Next U&
    '
    If REIMPU% > 0 Then
        Call COMUNI("Re-Imputación Completada")
      Else
        Call COMUNI("No se Pudo Completar la Re-Imputación")
    End If
    '
    GoTo 16
    '
99 Screen.MousePointer = 1
   Command1.Enabled = True
    '
End Sub

Sub Command10_Click()
  '\\\OT-06-0112-WAR-30/03/06///
  Command10.Enabled = False
    Call COMPENSACP
  Command10.Enabled = True
  '\\\OT-06-0112-WAR-FIN///
End Sub

Sub Command11_Click()
    Command11.Enabled = False
    OPX% = COMALTER%("Opciones de Consulta de Aplicaciones:\\Por Recibo o N/C\Por Factura o N/D\Listado por Rango de Fechas")
    If OPX% = 1 Then
        Call VAPLIREC
      ElseIf OPX% = 2 Then
        Call VAPLIFAC
      ElseIf OPX% = 3 Then
        Call LIAPLIFE
    End If
    Command11.Enabled = True
End Sub
'\\\OT-06-0409-MH-FIN///

Private Sub Command2_Click()
    '
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Hide
    '
End Sub

Sub Command20_Click()

   Command20.Enabled = False
   If InStr(1, EMPREAC$, "FERVI") > 0 Then
        '
10      OPXX% = ALTERNA%("Estadística Mensual de Cobranzas (Semanales)\Estadísticas de Cobranzas (Mensuales)\Recalculo Dias de Cobranza\Otros Accesos Directos")
        If OPXX% = 1 Then
           Call RECIBO08.ESTADISTICADECOBRANZA
           GoTo 10
         ElseIf OPXX% = 2 Then
           Call RECIBO08.ESTADISTICA_DE_COBRANZA_X_MES
           GoTo 10
         ElseIf OPXX% = 3 Then
           Call RECALCULO_DE_DIAS_DE_COBRANZA
           GoTo 10
         ElseIf OPXX% = 4 Then
           Call COMUNI("No Existen Accesos Directos \ para esta Aplicación")
           GoTo 10
         Else
           GoTo 99
         End If
    End If
    '
99  Command20.Enabled = True
    '
End Sub

Sub COBRANZACHEQPROPIO()

   Call DESHASFECHA(DES%, HAS%, PERIO$)
   If PERIO$ = "" Then GoTo 99
   DESDE = FETEXCOR1$(DES%)
   HASTA = FETEXCOR1$(HAS%)
   Call SELECHO("DEUDOR1")
   NCLI = XVALO(VALACT1$("DEUDOR1"))
   CUECAR% = CUEICARCH%
   If CUECAR% < 1 Then
        Call COMUNI("Imposible Emitir Reporte Falta Cuenta Contable de Cartera de Valores")
        GoTo 99
   End If
   '
   SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK) WHERE FECHCONTA>='" & DESDE & "' AND FECHCONTA<='" & HASTA & "' AND CUECONTAI=" & Str(CUECAR%)
   If NCLI > 0 Then SQLX$ = SQLX$ & " AND NUCLIEN=" & Str(NCLI)
   SQLX$ = SQLX$ & " AND NUCLIEN> 0 ORDER BY NUCLIEN ASC"
   Dim RST As ADODB.Recordset
   Set RST = READSET(SQLX$)
   '
   JJ& = 0
   With RST
        Do While Not .EOF
            NCLI = XVALO(!NUCLIEN)
            CUITCL$ = CuitCli$(NCLI)
            OBSER$ = SAFETEXT$(!OBSERGEN)
            If InStr(1, OBSER$, CUITCL$) > 0 Then
                FECHA% = CVINT(!fechconta)
                RECI$ = SAFETEXT$(!CodiCom_Cor)
                DATOSCHEQUE$ = SAFETEXT$(!DESCRIMOV)
                IMPORTE# = XVALO(!ValAbsComp)
                DOCSALIDA$ = SAFETEXT$(!Document)
                REFESALIDA$ = SAFETEXT$(!DESTICHEQ)
                '
                XX$ = REGBLAN$("COBCHPRO")
                
                Call REPLA(XX$, MKI$(NCLI), 1, 2)
                Call REPLA(XX$, MKI$(FECHA%), 3, 2)
                Call REPLA(XX$, RECI$, 5, 18)
                Call REPLA(XX$, DATOSCHEQUE$, 23, 48)
                Call REPLA(XX$, MKD$(IMPORTE#), 71, 8)
                Call REPLA(XX$, DOCSALIDA$, 79, 18)
                Call REPLA(XX$, REFESALIDA$, 97, 48)
                JJ& = JJ& + 1
                Call GRAREG("COBCHPRO", XX$, JJ&)
            End If
            .MoveNext
        Loop
   End With
   
   Call CONECRUN("*COBCHPRO", "Cobranza Con Cheques Propios")
   
99 Exit Sub
   
End Sub
 Sub Command23_Click()
    '
    Command23.Enabled = False
    If DERACCE%("REPIRECI", "Repetir Impresion de Recibo Emitido") >= 1 Then
       Call REPIRECIBO
    End If
    Command23.Enabled = True
    '
End Sub

 Sub Command25_Click()
    Command25.Enabled = False
    If DERACCE%("ANURECI", "Anulación de Recibos") >= 2 Then
       Call ANURECIBO
    End If
    Command25.Enabled = True
End Sub

 Sub Command3_Click()
    Command3.Enabled = False
    TC = TICAMONE(-1) ' fuerza refresco de tabla
    If DERACCE%("IMPUCRED", "Imputacion Créditos/Anticipos Deudores") >= 2 Then
      OPX% = COMALTER%("Opciones de Trabajo y Procesamiento:\\Imputar Créditos\Des-Aplicar Recibos\Cancelar")
      If OPX% = 1 Then
          On Error Resume Next
          Hide
          On Error GoTo 0
          Call FRESHALL
          Call IMPUCREVEN
          ' Call IMPUCREVEN_INCLUYE_AP_AUTOMATICA
        ElseIf OPX% = 2 Then
          Call DESAPLIREC
      End If
    End If
    Command3.Enabled = True
End Sub

 Sub Command4_Click()
    Command4.Enabled = False
    If DERACCE%("AJUSALD", "Ajuste de Saldos Deudores Clientes") >= 2 Then
       Call AJUSALCLI
    End If
    Command4.Enabled = True
End Sub

 Sub Command5_Click()
   Command5.Enabled = False
   Call COMUNI("Opción no Disponible en la Presente Versión.\  \Carga de Saldos Iniciales Clientes\por Ajuste de Saldos Deudores.")
   'Call CARSALINID
   Command5.Enabled = True
End Sub

Sub Command6_Click()
   Command6.Enabled = False
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EPAC$, "AHP") > 0 Then

10   OPX% = COMALTER%("Opciones\\Listado de Dif.Cambios\N.Debitos x Dif.Cambio")
     If OPX% < 1 Then GoTo 99
     If OPX% = 2 Then
        SOLNOTADEB.Show 1
        GoTo 10
     End If
   End If
   Call OPVARFA07.Command2_Click
99 Command6.Enabled = True
End Sub

Private Sub Command7_Click()
'    Call HELPONLINE("RECIBO")
Call ANUAJUSTE
End Sub

Sub CARSALINID()
   '
   If ULTREG&("SALINID") < 1 Then
      GoSub 4000
   End If
   '
2  x$ = REGLEIDO$("SALINID", 1)
   FEX = CVI(Mid$(x$, 9, 2))
   FSI% = FEX
   FAJU$ = FECHATEX$(FEX)
   '
3  cx% = ULTREG&("DEUDOR1")
   Dim SALINI#(32767), SUMDEU#(32767), SARRAS#(32767), REGAJU&(32767)
   '
   GoSub 6000    ' generar archivos eco de tipo y variante
   '
5  CLX$ = OBJPLA$("SELECLI", MKI$(NC%))
   If CLX$ = "" Then GoTo 200                   ' ajustar
   NC% = CVI(CLX$): If NC% < 1 Then Print Chr$(7);: GoTo 5
   CLI$ = rasocli$(NC%)
   '
   Call BLANARCH("!AJUSALD")
   '
6  Screen.MousePointer = 11
   FIN& = ULTREG&("SALINID")
   For i& = 1 To FIN&
      Call TRACE(20, i&, FIN&)
      x$ = REGLEIDO$("SALINID", i&)
      If CVI(Left$(x$, 2)) = NC% Then
          SIG% = 1
          TI% = Asc(Mid$(x$, 3, 1))
          VA% = Asc(Mid$(x$, 4, 1))
          If TI% < 0 Or TI% > 6 Then TI% = 2
          If TI% > 3 Then SIG% = (-1)
          Y$ = Mid$(x$, 3, 26) + MKS$(i&) + Mid$(x$, 31, 2)
          IPEN# = CVD(Mid$(x$, 11, 8)) * SIG%
          Call REPLA(Y$, MKD$(IPEN#), 9, 8)
          Call REGAPP("!AJUSALD", Y$)
      End If
   Next i&
   Screen.MousePointer = 1
   '
8  VTB% = VENTABU%("AJUSALD/NC/TITUPAN=" + CLI$)
   If VTB% < 0 Then GoTo 5
   ATX% = 0: ATY% = 0
   SALDEU# = 0
   '
   For U& = 1 To ULTREG&("!AJUSALD")
       '
       x$ = REGLEIDO$("!AJUSALD", U&)
       TI% = Asc(Left$(x$, 1))
       If TI% > 6 Then TI% = 2
       If TI% = 4 Then
          Call MENSERR(24, "Imposible Registrar Recibo\Como Saldo Inicial.\  \Registrar Como Ajuste de Saldo.")
          GoTo 8
       End If
       '
       If CVI(Mid$(x$, 17, 2)) > FSI% Then
          Call MENSERR(24, "Fecha de Origen Debe Ser Anterior\a Fecha Inicial.")
          GoTo 8
       End If
       '
       SIG% = 1: If TI% > 3 Then SIG% = (-1)
       SALEI# = CVD(Mid$(x$, 9, 8))
       SORI# = CVD(Mid$(x$, 19, 8))
       If Mid$(x$, 31, 2) <> String$(2, 255) Then
         If Abs(SALEI#) >= 0.005 Then
           If TI% > 0 Then
             If Sgn(SALEI#) <> SIG% Or Sgn(SORI#) <> Sgn(SALEI#) Then
               SALEI# = Abs(SALEI#) * SIG%
               Call REPLA(x$, MKD$(SALEI#), 9, 8)
               SORI# = (Abs(CVD(Mid$(x$, 19, 8)))) * Sgn(SALEI#)
               Call REPLA(x$, MKD$(SORI#), 19, 8)
               Call GRAREG("!AJUSALD", x$, U&)
               ATX% = 1
             End If
           End If
         End If
         ATY% = 1
       End If
       '
       SALDEU# = SALDEU# + SALEI#
   Next U&
   '
   If ATX% > 0 Then
       Call COMUNI("Errores de Signo de Saldos Deudores\han sido Corregidos Automáticamente.")
       GoTo 8
   End If
   '
100 FIN& = ULTREG&("!AJUSALD")
    Call BLOQARCH("SALINID")
    For i& = 1 To FIN&
      Call TRACE(20, i&, FIN&)
      x$ = REGLEIDO$("!AJUSALD", i&)
      REG& = CVS(Mid$(x$, 27, 4))
      '
      TI% = Asc(Left$(x$, 1))
      IPEN# = CVD(Mid$(x$, 9, 8)) * SIG%
      If TI% > 0 Then IPEN# = Abs(IPEN#)
      Y$ = MKI$(NC%) + Left$(x$, 26) + MKI$(0) + Mid$(x$, 31, 2)
      Call REPLA(Y$, MKD$(IPEN#), 11, 8)
      If REG& < 1 Then
           Call REGAPP("SALINID", Y$)
         Else
           If Mid$(Y$, 31, 2) <> String$(2, 255) Then
             Call GRAREG("SALINID", Y$, REG&)
           End If
      End If
   Next i&
   '
   GoTo 5
   '
   '
200 For CX1% = 1 To cx%
      SALINI#(CX1%) = 0
      SARRAS#(CX1%) = 0
      SUMDEU#(CX1%) = 0
    Next CX1%
    '
    Call BLOQARCH("SVD202")
    Call BLOQARCH("FACTUR")
    Call BLOQARCH("SALINID")
    FIN& = ULTREG&("SALINID")
    For i& = 2 To FIN&
      x$ = REGLEIDO$("SALINID", i&)
      If Len(x$) >= 32 Then
         NC% = CVI(Left$(x$, 2))
         If NC% > 0 Then
           SIG% = 1: TI% = Asc(Mid$(x$, 3, 1))
           VA% = Asc(Mid$(x$, 4, 1))
           If TI% < 0 Or TI% > 6 Then TI% = 2
           If TI% > 3 Then SIG% = (-1)
           '
           IMPO# = CVD(Mid$(x$, 11, 8)) * SIG%
           'If TI% > 3 Then IMPO# = (-1) * Abs(IMPO#)
           '
           If Mid$(x$, 31, 2) <> String$(2, 255) Then
             NRODO& = CVS(Mid$(x$, 5, 4))
             Call REPLA(x$, String$(2, 255), 31, 2)
             Call GRAREG("SALINID", x$, i&)
             If TI% > 0 Then
                Call REPLA(x$, MKD$(Abs(IMPO#)), 21, 8)
             End If
             Call REGAPP("SVD202", x$)
             Y$ = REGBLAN$("FACTUR")
             If TI% > 0 Then IMPO# = Abs(IMPO#)
             Call REPLA(Y$, MKI$(NC%) + MKS$(NRODO&) + MKI$(FSI%) + MKI$(FSI%) + String$(2, 0) + Chr$(TI%) + Chr$(VA%), 1, 14)
             Call REPLA(Y$, MKD$(IMPO#), 15, 8)
             Call REPLA(Y$, MKD$(IMPO#), 39, 8)
             Call REGAPP("FACTUR", Y$)
           End If
         End If
      End If
    Next i&
    '
    Call CIERRARCH("SALINID")
    Call CIERRARCH("FACTUR")
    Call CIERRARCH("SVD202")
    '
    Exit Sub
    '
1000 '************************************************ listar saldos iniciales
   '
   Call BLOQARCH("LISALIN")
   Call BLANARCH("LISALIN")
   '
   FIN& = ULTREG&("SALINID")
   For i& = 2 To FIN&
    Call TRACE(20, i&, FIN&)
    x$ = REGLEIDO$("SALINID", i&)
    NC% = CVI(Left$(x$, 2))
    If NC% > 0 Then
      TI% = Asc(Mid$(x$, 3, 1))
      TIX$ = "FC": SIG% = 1
          If TI% = 3 Then TIX$ = "ND"
          If TI% = 4 Then TIX$ = "RB": SIG% = (-1)
          If TI% = 6 Then TIX$ = "NC": SIG% = (-1)
      NRO& = CVS(Mid$(x$, 5, 4))
      NROX$ = TIX$ + AJUSTD$(TRIM$(Str$(NRO&)), 8)
      FF% = CVI(Mid$(x$, 19, 2))
      IORI# = Abs(CVD(Mid$(x$, 21, 8)))
      FVI% = CVI(Mid$(x$, 9, 2))
      IPEN# = CVD(Mid$(x$, 11, 8)) * SIG%
      DEBE# = IPEN#: Habe# = 0
      If Sgn(IPEN#) < 0 Then DEBE# = 0: Habe# = Abs(IPEN#)
      '
      Z$ = REGBLAN$("LISALIN")
      Call REPLA(Z$, MKI$(NC%), 1, 2)
      Call REPLA(Z$, NROX$, 3, 10)
      Call REPLA(Z$, MKI$(FF%), 13, 2)
      Call REPLA(Z$, MKD$(IORI#), 15, 8)
      Call REPLA(Z$, MKI$(FVI%), 23, 2)
      Call REPLA(Z$, MKD$(DEBE#), 25, 8)
      Call REPLA(Z$, MKD$(Habe#), 33, 8)
      Call REPLA(Z$, MKD$(0), 41, 8)
      '
      Call REGAPP("LISALIN", Z$)
      '
    End If
    '
   Next i&
   '
   Call CIERRARCH("LISALIN")
   Call FINAL("*SOSALIN")
   '

4000 '************************************** ingresa fecha de saldos iniciales
   '
   HO% = HOY(HOS$)
   Call SELECHO("SELFECHA")
   FAJU$ = TRIM$(VALACT1$("SELFECHA"))
   If FAJU$ = "" Then Exit Sub
   '
   FF% = FECHANUM(FAJU$)
   If FF% > HO% Then GoTo 4000
   '
   FSI% = FF%
   x$ = REGBLAN$("SALINID")
   x$ = REPLACE$(x$, MKI$(FSI%), 9, 2)
   x$ = REPLACE$(x$, MKI$(FSI%), 19, 2)
   Call GRAREG("SALINID", x$, 1)
   Call CIERRARCH("SALINID")
Return
   '

6000 '  ******************************* genera archivos eco de tipo y variante
   '
   If ULTREG&("ECOTIPO") <> 7 Then
       Call BLANARCH("ECOTIPO")
       Call REGAPP("ECOTIPO", MKI$(0) + "Ajuste de Saldo")
       Call REGAPP("ECOTIPO", MKI$(1) + "Venta Contado")
       Call REGAPP("ECOTIPO", MKI$(2) + "Factura C.Corr.")
       Call REGAPP("ECOTIPO", MKI$(3) + "Nota de Debito")
       Call REGAPP("ECOTIPO", MKI$(4) + "Anticipo")
       Call REGAPP("ECOTIPO", MKI$(5) + "Descuento Caja")
       Call REGAPP("ECOTIPO", MKI$(6) + "Nota de Credito")
       Call CIERRARCH("ECOTIPO")
   End If
   '
   If ULTREG&("ECOVARI") <> 4 Then
       Call BLANARCH("ECOVARI")
       Call REGAPP("ECOVARI", MKI$(0) + "No Definido")
       Call REGAPP("ECOVARI", MKI$(1) + "A - Responsables IVA")
       Call REGAPP("ECOVARI", MKI$(2) + "B - Consumidor Final")
       Call REGAPP("ECOVARI", MKI$(3) + "X - Exportacion")
       Call CIERRARCH("ECOVARI")
   End If
   '
   Return

End Sub

 Sub Command8_Click()
   '
   Command8.Enabled = False
   Dim RERECI&(10000)
   '
5  OPXX% = ALTERNA%("Avance Diario de Cobranza\Análisis Estadístico de Cobranza Por Cheques\Consulta de Cobranza con Tarjeta\Cobranza con Cheques Propios")
   If OPXX% < 1 Or OPXX% > 4 Then GoTo 99
   If OPXX% = 1 Then GoTo 10 ' Avance diario de cobranza
   If OPXX% = 2 Then Call ANACOCHEQ: GoTo 99
   If OPXX% = 3 Then Call FINAL("?COBCONTAR") ' consulta de Cobranza con Tarjeta
   If OPXX% = 4 Then Call COBRANZACHEQPROPIO ' consulta de Cobranza con Tarjeta
   GoTo 99
   '
10 FEX = HOY(HO$)
   HAS% = FEX
   Opciones% = COMALTER%("Selección de Fechas\\Por Día\Por Rango")
   If Opciones% < 1 Or Opciones% > 2 Then GoTo 5
   If Opciones% = 2 Then    ' POR RANGO
        If DERACCE%("COBPORFE", "Avance de Cobranza por Rango de Fechas") > 1 Then
             HASS$ = FECHATEX$(FEX)
             DESS$ = "01" + Mid$(HASS$, 3)
             DES% = FECHANUM(DESS$)
             VDEF$ = MKI$(DES%) + MKI$(HAS%)
             '
             OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
             If OBX$ = "" Then
                GoTo 10
             End If
             DES% = CVI(Left$(OBX$, 2))
             HAS% = CVI(Mid$(OBX$, 3, 2))
        Else
            GoTo 10
        End If
   Else                     ' POR DIA
        Call SELECHO("SELFECHA")
        FEAVACOB$ = TRIM$(VALACT1$("SELFECHA"))
        If FEAVACOB$ = "" Then GoTo 10
        '
        DES% = FECHANUM(FEAVACOB$)
        HAS% = DES%
   End If
   '
20 NUSUC% = 0
   If ULTREG&("COBRADOR") > 0 Then
      Call SELECHO("COBRADOR")
      TT2$ = VALACT2$("COBRADOR")
      If TRIM$(TT2$) = "" Then GoTo 10
      NUSUC% = XVALO(TRIM$(VALACT1$("COBRADOR")))
      If NUSUC% < 0 Then GoTo 10
   End If
   '
   Screen.MousePointer = 11
   '
   'DESA = Int(CDbl(CVDAT(DES%)))
   'HASA = Int(CDbl(CVDAT(HAS%)))
   
   'jandy sql
   
   DESA = FETEXCOR1$(DES%)
   HASA = FETEXCOR1$(HAS%)
   
   
   'nuevo
   Call APERBASDAT("CABAN")
   SQLX$ = "SELECT * FROM CAJABANC "
   SQLX$ = SQLX$ + "where FechConta >= '" & DESA & "'"
   SQLX$ = SQLX$ + "AND FechConta <= '" & HASA & "'"
   SQLX$ = SQLX$ + "AND NumeTalo = " & NUSUC%
   SQLX$ = SQLX$ + "AND TipoMovim = 'RCOB' "
   SQLX$ = SQLX$ + "Order by FechEmisi ASC "
   '\\\OT-06-0162-WAR-04/05/06///
   SQLX$ = SQLX$ + ", NumeComp ASC "
   '\\\OT-06-0162-WAR-FIN///
   'Set AVANCETMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
   
   Dim AVANCETMP As ADODB.Recordset
   Set AVANCETMP = New ADODB.Recordset
   AVANCETMP.CursorLocation = adUseClient
25 On Error Resume Next
   AVANCETMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   '
   With AVANCETMP
     On Error Resume Next
     .MoveFirst
     If Err Then
       On Error GoTo 0
       Call MENSERR(24, "Ningun Registro Coincide con el Filtro SQL")
       GoTo 99
     End If
     JJ& = 0
     REBLAY$ = REGBLAN$("AVAFACB")
     Do While Not .EOF
       NCLI = !NUCLIEN
       CLIEX$ = ""
       If NCLI > 0 Then
          CLIEX$ = rasocli$(NCLI)
          NRO& = !NumeComp
          NETO# = !ValAbsComp 'Abs((Int(100 * (CVD(Mid$(X$, 15, 8)) + CVD(Mid$(X$, 49, 8))) + 0.5)) / 100)
            If !OpciMovim = "NC" Then NETO# = (-1) * NETO#
          NUSUCUVE% = !CodiSucu
          NNX$ = !codicom_lar
          '
          MONEM% = 0: TICAM = 1
          On Error Resume Next
          MONEM% = !Mone_Emis
          On Error GoTo 0
          '
          TICAM = 1
          On Error Resume Next
          If IsNull(!TIPO_CAM) = False Then
              TICAM = !TIPO_CAM
          End If
          On Error GoTo 0
          '
          Y$ = REBLAY$
          Call REPLA(Y$, MKI$(CVINT(!fechconta)), 1, 2)
          Call REPLA(Y$, CLIEX$, 3, 28)
          Call REPLA(Y$, NNX$, 31, 18)
          Call REPLA(Y$, MKD$(NETO#), 49, 8)
          '
          If MONEM% > 1 Then
             If TICAM > 0.1 Then
                Call REPLA(Y$, SIMBOLPES$(MONEM%), 65, 4)
                Call REPLA(Y$, MKD$(NETO# / TICAM), 69, 8)
             End If
          End If
          '
          JJ& = JJ& + 1
          Call GRAREG("AVAFACB", Y$, JJ&)
       End If
49      .MoveNext
     Loop
   End With
   AVANCETMP.Close
   Set AVANCETMP = Nothing
   '
   Call SETULTREG("AVAFACB", JJ&)
   Call CIERRARCH("*.*")
   '
   If JJ& < 1 Then
     Call MENSERR(24, "Sin Movimientos\en el Rango de Fechas Elegido")
     GoTo 10
   End If
   '
   Call CIERRARCH("*.*")
   Screen.MousePointer = 1
   '
   Call CONECRUN("*LIAVAFB", "Secuencia de Recibos Registrados/WAIT")
   If ULTREG&("COBRADOR") > 1 Then
     GoTo 20
   End If
   '
99 Screen.MousePointer = 1
   Command8.Enabled = True
   '
End Sub

Sub ANACOCHEQ()
    '
    Dim INIPER%(12), FINPER%(12)
' aquí poner un deshasfecha y obtener des%,HAS%,PERIO$
''    MMXX0$ = TRIM$(MESAN.TEXT)
''    MMXX1$ = MESANO$(MMXX0$, DES%, HAS%)
''    If TRIM$(MMXX1$) = "" Then
''        Call MENSERR(24, "Ingrese el Mes a Listar")
''        Exit Sub
''    End If

    
    Call CARSELMESA
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) = "" Then
        Exit Sub
    End If
    '
    MMXX0$ = TRIM$(MESASEL$)
    MMXX1$ = MESANO$(MMXX0$, DES%, HAS%)
    '
    DESA = FETEXCOR1$(DES%)
    HASA = FETEXCOR1$(HAS%)
    
    '
' ESTO SE SUPRIME ... VA AUTOMATICAMENTE A IMPRESION
''    If Option11.Value = True Then
''        Call MENSERR(24, "Imposible por Pantalla.\Exporte a Archivo de Texto.")
''        Exit Sub
''    End If
    '
    REAVA& = 0
    Screen.MousePointer = 11
    '
    DD1% = DES%
    FEX = DD1%
    SMX0% = SEMANU%(FEX, DELX, ALX)
    INIPER%(0) = DELX: FINPER%(0) = ALX
    For KKI% = 1 To 5
      FEX = DIPOST(ALX)
      SMX1% = SEMANU%(FEX, DELX, ALX)
      INIPER%(KKI%) = DELX: FINPER%(KKI%) = ALX
    Next KKI%
    '
    FEX = INIPER%(3)
    FETE$ = FECHATEX$(FEX)
    MACT% = 1 + Val(Mid$(FETE$, 4, 2))
    If MACT% > 12 Then MACT% = 1
    '
    FEX = FINPER%(4)
    INIPER%(5) = DIPOST(FEX)
    While Val(Mid$(FECHATEX$(FEX), 4, 2)) <> MACT%
      FEX = DIPOST(FEX)
    Wend
10  FEY = DIPOST(FEX)
    If Val(Mid$(FECHATEX$(FEY), 4, 2)) = MACT% Then
       FEX = FEY
       GoTo 10
    End If
    FINPER%(5) = FEX
    '
    MACT% = MACT + 1
    If MACT% > 12 Then MACT% = 1
    '
    FEX = FINPER%(5)
    INIPER%(6) = DIPOST(FEX)
    While Val(Mid$(FECHATEX$(FEX), 4, 2)) <> MACT%
      FEX = DIPOST(FEX)
    Wend
20  FEY = DIPOST(FEX)
    If Val(Mid$(FECHATEX$(FEY), 4, 2)) = MACT% Then
       FEX = FEY
       GoTo 20
    End If
    FINPER%(6) = FEX
    '
    MACT% = MACT + 1
    If MACT% > 12 Then MACT% = 1
    '
    FEX = FINPER%(6)
    INIPER%(7) = DIPOST(FEX)
    While Val(Mid$(FECHATEX$(FEX), 4, 2)) <> MACT%
      FEX = DIPOST(FEX)
    Wend
30  FEY = DIPOST(FEX)
    If Val(Mid$(FECHATEX$(FEY), 4, 2)) = MACT% Then
       FEX = FEY
       GoTo 30
    End If
    FINPER%(7) = 32767
    '
''    JJ& = 0
''    REBLA$ = REGBLAN$("ANACHERE")
''' AQUI HACER UNA SELECT COBRE CAJABANC, CON TIPOMOVIM = 'CHRECI', DESDE Y HASTA FECHA DE RECEPCION Y CUECONTAI% = CUENTA DE CARTERA DE VALORES.
''    For U& = 1 To ULTREG&("CHECAR")   ' REEMPLAZAR POR UN DO WHILE NOT .EOF
''      X$ = REGLEIDO$("CHECAR", U&)    ' CONSIDERAR EL RECORDSET
''      FECOBRA% = CVI(Mid$(X$, 45, 2)) ' CONSIDERAR EL RECORDSER CON CVINT
''      If FECOBRA% >= DES% Then        ' NO HACE FALTA POR EL FILTRO EN EL RECORDSET
''        If FECOBRA% <= HAS% Then      ' NO HACE FALTA POR EL FILTRO EN EL RECORDSET
''          NURECI = CVS(Mid$(X$, 47, 4))   ' VALOR DEL RECORDSET
''          NCLI = CVI(Mid$(X$, 3, 2))    ' VALOR DEL RECORDSET
''          BANCHE$ = TRIM$(Mid$(X$, 17, 16))     ' VALOR DEL RECORDSET
''          If BANCHE$ = "" Then BANCHE$ = TRIM$(ECOARCH$("MECOBRA", Left$(X$, 2)))
''          NUCHEBA$ = BANCHE$
''          If NUCHEBA$ <> "" And TRIM$(Mid$(X$, 5, 12)) <> "" Then NUCHEBA$ = NUCHEBA$ + " - "
''          NUCHEBA$ = NUCHEBA$ + TRIM$(Mid$(X$, 5, 12))
''          IMPOCHE# = CVD(Mid$(X$, 33, 8))      ' VALOR DEL RECORDSET
''          FACRE% = CVI(Mid$(X$, 41, 2))        ' VALOR DEL RECORDSET
''          If FACRE% < INIPER%(0) Then FACRE% = INIPER%(0)
          
    JJ& = 0
    i& = 0
    sqlx1$ = "TIPOMOVIM = 'CHRECI' "
    sqlx1$ = sqlx1$ + "AND FECHCONTA >= '" & DESA & "' "
    sqlx1$ = sqlx1$ + " AND FECHCONTA <= '" & HASA & "' "
    
    FIN& = CantRegistros&("CAJABANC", sqlx1$)
    
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + " WHERE " & sqlx1$
    SQLX$ = SQLX$ + " ORDER BY FechConta,NUMECOMP ASC"

    '
    Dim CABATMP As ADODB.Recordset
    Set CABATMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    With CABATMP
      Do While Not .EOF
       Call TRACE(20, i&, FIN&)
       i& = i& + 1
       FECOBRA% = CVINT(!fechconta)
       NCLI = XVALO(!NUCLIEN)
       NURECI = XVALO(!NumeComp)
       REFERE$ = SAFETEXT$(!DESCRIMOV)
       BANCHE$ = TRIM$(Left$(REFERE$, 18))
       'If BANCHE$ = "" Then BANCHE$ = TRIM$(ECOARCH$("MECOBRA", Left$(X$, 2))) ** COMENTADO P0R QUE NO TENGO EL MEDIO DE COBRANZA PARA PASAR
       NUCHEBA$ = TRIM$(Mid$(REFERE$, 19, 14))
       If NUCHEBA$ <> "" And TRIM$(Mid$(REFERE$, 19, 14)) <> "" Then NUCHEBA$ = NUCHEBA$ + " - "
       NUCHEBA$ = NUCHEBA$ + BANCHE$
       'If FACRE% < INIPER%(0) Then FACRE% = INIPER%(0)
       IMPOCHE# = XVALO(!idebecomp)
       FACRE% = CVINT(!fechacred)
       If FACRE% < INIPER%(0) Then FACRE% = INIPER%(0)
       
         
          '
 ' ESTO QUE SIGUE QUEDA IGUAL
       Y$ = REBLA$
       Call REPLA(Y$, MKS$(NURECI), 1, 4)
       Call REPLA(Y$, MKI$(FECOBRA%), 5, 2)
       Call REPLA(Y$, MKS$(NCLI), 7, 4)
       Call REPLA(Y$, NUCHEBA$, 11, 32)
       Call REPLA(Y$, MKI$(FACRE%), 43, 2)
       For KKI% = 0 To 7
            If FACRE% <= FINPER%(KKI%) Then
              Call REPLA(Y$, MKD$(IMPOCHE#), 45 + 8 * KKI%, 8)
              GoTo 40
            End If
       Next KKI%
40     Call REPLA(Y$, MKD$(IMPOCHE#), 109, 8)
       JJ& = JJ& + 1
       Call GRAREG("ANACHERE", Y$, JJ&)
      .MoveNext
     Loop
    End With

    '
    Call SETULTREG("ANACHERE", JJ&)
99  Call CIERRARCH("*.*")
    Call HEADERS("ANACHERE", "")
' EL MMXX1 ES EL PERIO$
    Call HEADERS("ANACHERE", "Corresponde a: " + MMXX1$)
' DEBERIA FUNCIONAR
    Call FINAL("*LIANACHE")
    Screen.MousePointer = 1
   '
End Sub

Sub ANURECIBO()
    '
    If DERACCE%("ANURECO", "Anulación de Recibos de Cobranza") < 2 Then
      Exit Sub
    End If
    '
    ' 1.- Ingresar Numero de Cliente
    '     ****************************
    '
    FECHALIM% = XVALO(Control$("", "CIEMEVEN")) '\\\OT-06-0212-WAR-06/06/06///
    'jandy sql
    'FECHALU = Int(CDbl(CVDAT(FECHALIM%)))
    FECHALU = FECHATEX(FECHALIM%)
    '
'10  Call SELECHO("DEUDOR1")
'    NCLIE = XVALO(VALACT1$("DEUDOR1"))
10  NCLIE = ListaSeleccionClientes()
    If NCLIE < 1 Then Exit Sub
    '
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + " WHERE (TipoMovim = 'RCOB' OR (TipoMovim = 'AJUD' AND OpciMovim = 'CC')) "
    SQLX$ = SQLX$ + " AND NuCLIEN = " & NCLIE & " "
    'SQLX$ = SQLX$ + "AND FechEmisi > " & FECHALU & " " '\\\OT-06-0212-WAR-06/06/06///
    SQLX$ = SQLX$ + " AND FechConta > ' " & FECHALU & "' "
    SQLX$ = SQLX$ + " Order by CODICOM_LAR ASC "
    '
    'Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim CABATEMP As ADODB.Recordset
    Set CABATEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    JJ& = 0
    With CABATEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        Do While Not .EOF
          IMPO# = !ValAbsComp
          If IMPO# >= 0.005 Then
            IMPOS$ = NUMSTRI$(IMPO#, 14, 2, 1)
            NPX& = XVALO(Mid$(!codicom_lar, 11, 8))
            JJ& = JJ& + 1
            'Y$ = MKS$(NPX&) + AJUSTI$(!codicom_lar, 18) + IMPOS$
            Y$ = MKS$(JJ&) + AJUSTI$(!codicom_lar, 18) + IMPOS$
            Call GRAREG("!SELCUERE", Y$, JJ&)
          End If
        .MoveNext
        Loop
      End If
    End With
    CABATEMP.Close
    Set CABATEMP = Nothing
    
    Call SETULTREG("!SELCUERE", JJ&)
    Screen.MousePointer = 1
    '
    If ULTREG&("!SELCUERE") < 1 Then
       Call MENSERR(24, "La Cuenta Elegida no tiene Comprobantes de Cobranza \ que Puedan Anularse.")
       GoTo 10
    End If
    '
    Call CIERRARCH("!SELCUERE")
    Call FRESHECHO("!SELCUERE")
    '
    Call SELECHO("!SELCUERE")
    If TRIM$(VALACT2$("!SELCUERE")) = "" Then GoTo 10
    '
    DOCULA$ = TRIM$(Left$(VALACT2$("!SELCUERE"), 18))
    DOCULb$ = "AS" + TRIM(Mid(DOCULA$, 3, 18))
    NUOPANU& = XVALO(Mid$(DOCULA$, 13, 6))
    '
    '\\\OT-05-0736-WAR-11/11/05/// NO ANULA SI CHEQUE ENTREGADO O DEPOSITADO
    Call APERBASDAT("CABAN")
    Dim CABANSEL As ADODB.Recordset
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE (CodiCom_Lar = '" & DOCULA$ & "' OR CodiCom_Lar = '" & DOCULb$ & "') "
    SQLX$ = SQLX$ + "AND NuCLIEN = " & NCLIE & " "
    'SQLX$ = SQLX$ + "AND FechEmisi > " & FECHALU & " " '\\\OT-06-0212-WAR-06/06/06///
    SQLX$ = SQLX$ + "AND FechConta > '" & FECHALU & "' "
    'Set CABANSEL = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
    Set CABANSEL = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With CABANSEL
      Do While Not (.EOF)
        If !codicom_lar <> DOCULA$ Then GoTo 17
        
      'SQLX$ = "CodiCom_Lar = '" & DOCULA$ & "' AND NuCLIEN = " & NCLIE & " "
      '.FindFirst SQLX$
'15    If .NoMatch = False Then
        If UCase$(!OpciMovim) <> "CH" Then GoTo 17 ' NO ES MOVIMIENTO DE CHEQUE
        If !Status > 1 Then GoTo 16               ' SALIDO DE CARTERA
        If IsNull(!FECHENTRE) = True Then GoTo 17  ' NO TIENE FECHA DE ENTREGA
        If CVINT(!FECHENTRE) < 33 Then GoTo 17     ' LA FECHA DE ENTREGA ES MENOR A 33
        '
16      Call MENSERR(24, "Anulación Imposible.\ Valores Ingresados con este Recibo\Figuran Depositados o Entregados en Pago.")
        GoTo 999
        '
17      '.FindNext SQLX$
        .MoveNext
        Loop
      'End If
    End With
    CABANSEL.Close
    Set CABANSEL = Nothing
    
    '\\\OT-05-0736-WAR-FIN///
    '
18  RESERNUM% = COMALTER%("Desea Dejar Reservado\este Número de Comprobante ?\\Reservar Número\Re-Utilizarlo")
    If RESERNUM% < 1 Or RESERNUM% > 2 Then GoTo 18
    '
'    SQLX$ = "SELECT * FROM APLICOBRA "
'    SQLX$ = SQLX$ + "WHERE (CoDocOr_Lar = '" & DOCULA$ & "' OR CoDocAp_Lar = '" & DOCULA$ & "' "
'    SQLX$ = SQLX$ + "OR CoDocOr_Lar = '" & DOCULb$ & "' OR CoDocAp_Lar = '" & DOCULb$ & "') "
'    SQLX$ = SQLX$ + "AND NuCLIEN = " & NCLIE & " "
'    Set APLICSEL = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
'    '
'    wrkPredeterminado.BeginTrans
'    With CABANSEL
'      .FindFirst "(CodiCom_Lar = '" & DOCULA$ & "' OR CodiCom_Lar = '" & DOCULB$ & "') AND NuCLIEN = " & NCLIE & " "
'21    If .NoMatch = False Then
'        If RESERNUM% <> 1 Or TRIM$(!TipoMovim) <> "RCOB" Then
'             .Delete
'           Else
'             .Edit
'             !ValAbsComp = 0
'             !IDEBECOMP = 0
'             !IHABECOMP = 0
'             !BASEIMPO = 0
'             !IdenBaja = IDACTU$
'             !FeHoReal = Now
'             .Update
'        End If
'        .FindNext "(CodiCom_Lar = '" & DOCULA$ & "' OR CodiCom_Lar = '" & DOCULB$ & "') AND NuCLIEN = " & NCLIE & " "
'        GoTo 21
'      End If
'    End With
'
'    With APLICSEL
'31    .FindFirst "NuClien = " & NCLIE & " "
'      If .NoMatch = False Then
'         .Delete
'         GoTo 31
'      End If
'    End With
    
    'jandy sql lo hice junto con ernesto
    
    strSQL = "Delete Cajabanc "
    strSQL = strSQL & " WHERE (CodiCom_Lar = '" & DOCULA$ & "' OR CodiCom_Lar = '" & DOCULb$ & "') "
    strSQL = strSQL & " AND NuCLIEN = " & NCLIE
    If RESERNUM% = 1 Then
       strSQL = strSQL & " And tipomovim <>'RCOB' "
    End If
    Call EJECUTACOMANDO(strSQL)
    '
 ' ojo - ver de recuperar que grabe el idactu$ !!!!!!!!!!!!!!!!!!!!!!!
    '
    strSQL = "Update Cajabanc set ValAbsComp = 0 "
    strSQL = strSQL & ", IDEBECOMP = 0, IHABECOMP = 0"
    strSQL = strSQL & ", BASEIMPO = 0 "        ' , IdenBaja = '" & IDACTU$ & "' "
    strSQL = strSQL & ", FeHoReal=getdate()"
    strSQL = strSQL & " WHERE (CodiCom_Lar = '" & DOCULA$ & "' OR CodiCom_Lar = '" & DOCULb$ & "') "
    strSQL = strSQL & " AND NuCLIEN = " & NCLIE
    Call EJECUTACOMANDO(strSQL)
   
    strSQL = "Delete from APLICOBRA "
    strSQL = strSQL & " WHERE (CoDocOr_Lar = '" & DOCULA$ & "' OR CoDocAp_Lar = '" & DOCULA$ & "' "
    strSQL = strSQL & " OR CoDocOr_Lar = '" & DOCULb$ & "' OR CoDocAp_Lar = '" & DOCULb$ & "') "
    strSQL = strSQL & " AND NuCLIEN = " & NCLIE & " "
    Call EJECUTACOMANDO(strSQL)
    '
    Condi$ = " NuClien = " & NCLIE & " AND TIPOMOVIM = 'AjReci'  and left(descrimov ,18) =  '" & DOCULA$ & "' "
    Call BORRAREGISTROS("CAJABANC", Condi$, REGAF&)

'    RFG$ = RECFILT$("CARBIMON", 2, MKI$(NCLIE))
'    For KUL& = Len(RFG$) - 3 To 1 Step -4
'      RECABIN& = CVS(Mid$(RFG$, KUL&, 4))
'      XXC$ = REGLEIDO$("CARBIMON", RECABIN&)
'      If (TRIM$(Mid$(XXC$, 45, 18)) = TRIM$(DOCULA$) Or TRIM$(Mid$(XXC$, 45, 18)) = TRIM$(DOCULb$)) Then
'        DOCABE$ = TRIM$(Mid$(XXC$, 13, 18))
'        Call GRAREG("CARBIMON", String$(256, 0), RECABIN&)
'      End If
'    Next KUL&
    '
    Call COMUNI("Anulación Comprobante " + TRIM$(DOCULA$) + " Completa")
    '
    'RECUEC07.Text1.TEXT = ""
    'RECUEC07.Text1.TEXT = TRIM$(Str$(NCLIE))
    Call RECUEC07.REAPLICRE(NCLIE)
    On Error Resume Next
    Unload RECUEC07
    Hide
    On Error GoTo 0
    Call FRESHALL
    PROLLAMA$ = Name
    'RECUEC07.Show 1
    Call FRESHALL
    '
    '
999 Call CIERRARCH("*.*")
    Call LIBARCH("!SELCUERE")
    Screen.MousePointer = 1
    '
    GoTo 10
    '
End Sub
Sub ANUAJUSTE()
    '
    If DERACCE%("ANURECO", "Anulación de Recibos de Cobranza") < 2 Then
      Exit Sub
    End If
    '
    ' 1.- Ingresar Numero de Cliente
    '     ****************************
    '
    FECHALIM% = XVALO(Control$("", "CIEMEVEN")) '\\\OT-06-0212-WAR-06/06/06///
    'jandy sql
    'FECHALU = Int(CDbl(CVDAT(FECHALIM%)))
    FECHALU = FECHATEX(FECHALIM%)
    '
'10  Call SELECHO("DEUDOR1")
'    NCLIE = XVALO(VALACT1$("DEUDOR1"))
10  NCLIE = ListaSeleccionClientes()
    If NCLIE < 1 Then Exit Sub
    '
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + " WHERE (TipoMovim = 'AJUD') "
    'SQLX$ = SQLX$ + "AND FechEmisi > " & FECHALU & " " '\\\OT-06-0212-WAR-06/06/06///
    'SQLX$ = SQLX$ + " AND FechEmisi > ' " & FECHALU & "' "
    SQLX$ = SQLX$ + " Order by CODICOM_LAR ASC "
    '
    'Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim CABATEMP As ADODB.Recordset
    Set CABATEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    JJ& = 0
    With CABATEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        Do While Not .EOF
          IMPO# = !ValAbsComp
          If IMPO# >= 0.005 Then
            IMPOS$ = NUMSTRI$(IMPO#, 14, 2, 1)
            NPX& = XVALO(Mid$(!codicom_lar, 11, 8))
            JJ& = JJ& + 1
            'Y$ = MKS$(NPX&) + AJUSTI$(!codicom_lar, 18) + IMPOS$
            Y$ = MKS$(JJ&) + AJUSTI$(!codicom_lar, 18) + IMPOS$
            Call GRAREG("!SELCUERE", Y$, JJ&)
          End If
        .MoveNext
        Loop
      End If
    End With
    CABATEMP.Close
    Set CABATEMP = Nothing
    
    Call SETULTREG("!SELCUERE", JJ&)
    Screen.MousePointer = 1
    '
    If ULTREG&("!SELCUERE") < 1 Then
       Call MENSERR(24, "La Cuenta Elegida no tiene Comprobantes de Cobranza \ que Puedan Anularse.")
       GoTo 10
    End If
    '
    Call CIERRARCH("!SELCUERE")
    Call FRESHECHO("!SELCUERE")
    '
    Call SELECHO("!SELCUERE")
    If TRIM$(VALACT2$("!SELCUERE")) = "" Then GoTo 10
    '
    DOCULA$ = TRIM$(Left$(VALACT2$("!SELCUERE"), 18))
    DOCULb$ = "AS" + TRIM(Mid(DOCULA$, 3, 18))
    NUOPANU& = XVALO(Mid$(DOCULA$, 13, 6))
    '
    '\\\OT-05-0736-WAR-11/11/05/// NO ANULA SI CHEQUE ENTREGADO O DEPOSITADO
    Call APERBASDAT("CABAN")
    Dim CABANSEL As ADODB.Recordset
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE (CodiCom_Lar = '" & DOCULA$ & "' OR CodiCom_Lar = '" & DOCULb$ & "') "
    SQLX$ = SQLX$ + "AND NuCLIEN = " & NCLIE & " "
    'SQLX$ = SQLX$ + "AND FechEmisi > " & FECHALU & " " '\\\OT-06-0212-WAR-06/06/06///
    'SQLX$ = SQLX$ + "AND FechEmisi > '" & FECHALU & "' "
    'Set CABANSEL = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
    Set CABANSEL = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With CABANSEL
      Do While Not (.EOF)
        If !codicom_lar <> DOCULA$ Then GoTo 17
        
      'SQLX$ = "CodiCom_Lar = '" & DOCULA$ & "' AND NuCLIEN = " & NCLIE & " "
      '.FindFirst SQLX$
'15    If .NoMatch = False Then
        If UCase$(!OpciMovim) <> "CH" Then GoTo 17 ' NO ES MOVIMIENTO DE CHEQUE
        If !Status > 1 Then GoTo 16               ' SALIDO DE CARTERA
        If IsNull(!FECHENTRE) = True Then GoTo 17  ' NO TIENE FECHA DE ENTREGA
        If CVINT(!FECHENTRE) < 33 Then GoTo 17     ' LA FECHA DE ENTREGA ES MENOR A 33
        '
16      Call MENSERR(24, "Anulación Imposible.\ Valores Ingresados con este Recibo\Figuran Depositados o Entregados en Pago.")
        GoTo 999
        '
17      '.FindNext SQLX$
        .MoveNext
        Loop
      'End If
    End With
    CABANSEL.Close
    Set CABANSEL = Nothing
    
    '\\\OT-05-0736-WAR-FIN///
    '
18  RESERNUM% = COMALTER%("Desea Dejar Reservado\este Número de Comprobante ?\\Reservar Número\Re-Utilizarlo")
    If RESERNUM% < 1 Or RESERNUM% > 2 Then GoTo 18
    '
'    SQLX$ = "SELECT * FROM APLICOBRA "
'    SQLX$ = SQLX$ + "WHERE (CoDocOr_Lar = '" & DOCULA$ & "' OR CoDocAp_Lar = '" & DOCULA$ & "' "
'    SQLX$ = SQLX$ + "OR CoDocOr_Lar = '" & DOCULb$ & "' OR CoDocAp_Lar = '" & DOCULb$ & "') "
'    SQLX$ = SQLX$ + "AND NuCLIEN = " & NCLIE & " "
'    Set APLICSEL = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
'    '
'    wrkPredeterminado.BeginTrans
'    With CABANSEL
'      .FindFirst "(CodiCom_Lar = '" & DOCULA$ & "' OR CodiCom_Lar = '" & DOCULB$ & "') AND NuCLIEN = " & NCLIE & " "
'21    If .NoMatch = False Then
'        If RESERNUM% <> 1 Or TRIM$(!TipoMovim) <> "RCOB" Then
'             .Delete
'           Else
'             .Edit
'             !ValAbsComp = 0
'             !IDEBECOMP = 0
'             !IHABECOMP = 0
'             !BASEIMPO = 0
'             !IdenBaja = IDACTU$
'             !FeHoReal = Now
'             .Update
'        End If
'        .FindNext "(CodiCom_Lar = '" & DOCULA$ & "' OR CodiCom_Lar = '" & DOCULB$ & "') AND NuCLIEN = " & NCLIE & " "
'        GoTo 21
'      End If
'    End With
'
'    With APLICSEL
'31    .FindFirst "NuClien = " & NCLIE & " "
'      If .NoMatch = False Then
'         .Delete
'         GoTo 31
'      End If
'    End With
    
    'jandy sql lo hice junto con ernesto
    
    strSQL = "Delete Cajabanc "
    strSQL = strSQL & " WHERE (CodiCom_Lar = '" & DOCULA$ & "' OR CodiCom_Lar = '" & DOCULb$ & "') "
    strSQL = strSQL & " AND NuCLIEN = " & NCLIE
    If RESERNUM% = 1 Then
       strSQL = strSQL & " And tipomovim <>'AJUD' "
    End If
    Call EJECUTACOMANDO(strSQL)
    '
 ' ojo - ver de recuperar que grabe el idactu$ !!!!!!!!!!!!!!!!!!!!!!!
    '
    strSQL = "Update Cajabanc set ValAbsComp = 0 "
    strSQL = strSQL & ", IDEBECOMP = 0, IHABECOMP = 0"
    strSQL = strSQL & ", BASEIMPO = 0 "        ' , IdenBaja = '" & IDACTU$ & "' "
    strSQL = strSQL & ", FeHoReal=getdate()"
    strSQL = strSQL & " WHERE (CodiCom_Lar = '" & DOCULA$ & "' OR CodiCom_Lar = '" & DOCULb$ & "') "
    strSQL = strSQL & " AND NuCLIEN = " & NCLIE
    Call EJECUTACOMANDO(strSQL)
   
    strSQL = "Delete from APLICOBRA "
    strSQL = strSQL & " WHERE (CoDocOr_Lar = '" & DOCULA$ & "' OR CoDocAp_Lar = '" & DOCULA$ & "' "
    strSQL = strSQL & " OR CoDocOr_Lar = '" & DOCULb$ & "' OR CoDocAp_Lar = '" & DOCULb$ & "') "
    strSQL = strSQL & " AND NuCLIEN = " & NCLIE & " "
    Call EJECUTACOMANDO(strSQL)
    '
    Condi$ = " NuClien = " & NCLIE & " AND TIPOMOVIM = 'AjReci'  and left(descrimov ,18) =  '" & DOCULA$ & "' "
    Call BORRAREGISTROS("CAJABANC", Condi$, REGAF&)

'    RFG$ = RECFILT$("CARBIMON", 2, MKI$(NCLIE))
'    For KUL& = Len(RFG$) - 3 To 1 Step -4
'      RECABIN& = CVS(Mid$(RFG$, KUL&, 4))
'      XXC$ = REGLEIDO$("CARBIMON", RECABIN&)
'      If (TRIM$(Mid$(XXC$, 45, 18)) = TRIM$(DOCULA$) Or TRIM$(Mid$(XXC$, 45, 18)) = TRIM$(DOCULb$)) Then
'        DOCABE$ = TRIM$(Mid$(XXC$, 13, 18))
'        Call GRAREG("CARBIMON", String$(256, 0), RECABIN&)
'      End If
'    Next KUL&
    '
    Call COMUNI("Anulación Comprobante " + TRIM$(DOCULA$) + " Completa")
    '
    'RECUEC07.Text1.TEXT = ""
    'RECUEC07.Text1.TEXT = TRIM$(Str$(NCLIE))
    Call RECUEC07.REAPLICRE(NCLIE)
    On Error Resume Next
    Unload RECUEC07
    Hide
    On Error GoTo 0
    Call FRESHALL
    PROLLAMA$ = Name
    'RECUEC07.Show 1
    Call FRESHALL
    '
    '
999 Call CIERRARCH("*.*")
    Call LIBARCH("!SELCUERE")
    Screen.MousePointer = 1
    '
    GoTo 10
    '
End Sub


Sub AJUSALCLI()
    '
    If MODOMONEI% = 2 Then
      OPXX% = ALTERNA%("Ajuste Cuenta Corriente\Saldo Factura Moneda Extranjera")
      If OPXX% = 1 Then
          GoTo 3
        ElseIf OPXX% = 2 Then
          Call AJUSALDOL
          Exit Sub
        Else
          Exit Sub
      End If
    End If
    '
3   Dim FVIK%(10), SD#(10), IMC#(10)
    HOII% = HOY(HO$)
    FAJUI% = HOII%
    '
5   Screen.MousePointer = 1
    VDEF$ = MKS$(NC%) + MKI$(FAJUI%) 'TOMA POR DEFECTO LA MONEDA PESOS
    CLSS$ = OBJPLA$("SELCLIAJUN", VDEF$)
    
    If CLSS$ = "" Or CLSS$ = "<ESC>" Then GoTo 99
    NCLI = CVS(CLSS$)
    If NCLI < 1 Then GoTo 5
    CLI$ = rasocli$(NCLI)
    FAJUI% = CVI(Mid$(CLSS$, 5, 2))
    If FAJUI% < 1 Or FAJUI% > HOII% Then GoTo 5
    '
    FECHALIM% = XVALO(Control$("", "CIEMEVEN"))
    If FAJUI% <= FECHALIM% Then
        Call MENSERR(24, "Fecha no Válida.\  \Ese Período Mensual Está Cerrado.")
        GoTo 5
    End If
    '
    Screen.MousePointer = 11
    '
    SARRASFE# = 0: SALFINAL# = 0: SALDEUTO# = 0
    '
'2.- Determina Saldo por arrastre de Facturas y Debitos
    
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "where NuClien = " & NCLI & " "
    SQLX$ = SQLX$ + "AND (TipoMovim = 'FVEN' "
    SQLX$ = SQLX$ + "or TipoMovim = 'RCOB' "
    SQLX$ = SQLX$ + "or TipoMovim = 'AJUD') "
    '
    'Set CABATEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    'jandy sql se puede hacer con un group by
    Dim CABATEMP As ADODB.Recordset
    Set CABATEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With CABATEMP
      On Error Resume Next
      .MoveFirst
      If Err > 0 Then
        On Error GoTo 0
        GoTo 100
      End If
      On Error GoTo 0
      Do While Not .EOF
        IFAC# = ROUND#(!idebecomp - !ihabecomp)
        FF% = CVINT(!fechconta)
        SALFINAL# = SALFINAL# + IFAC#
        If FF% <= FAJUI% Then
          SARRASFE# = SARRASFE# + IFAC#
        End If
      .MoveNext
      Loop
    End With
    CABATEMP.Close
    Set CABATEMP = Nothing
    '
'3.- Determina saldo por composicion de deuda
'    SQLX$ = "SELECT * FROM SADEUPEN "
'    SQLX$ = SQLX$ + "where NuClien = " & NC% & " "
'    '
    'Set CODEUTEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    
    strSQL = "Select Sum(SaldDebe - SALDAcre) as SALDEUTO from SaDeuPen"
    strSQL = strSQL & " where NuClien = " & NCLI
    Dim MIRS As ADODB.Recordset
    Set MIRS = FLEXCONN.Execute(FILTSQL$(strSQL))
    If Not (MIRS.EOF And MIRS.BOF) Then
      SALDEUTO# = ROUND#(MIRS!SALDEUTO)
    End If
    MIRS.Close
    Set MIRS = Nothing

'    With CODEUTEMP
'      On Error Resume Next
'      .MoveFirst
'      If Err > 0 Then
'        On Error GoTo 0
'        GoTo 100
'      End If
'      On Error GoTo 0
'      Do While Not .EOF
'        SALDEUTO# = SALDEUTO# + ROUND#(!SaldDebe - !SALDAcre)
'      .MoveNext
'      Loop
'    End With
 
    '
100 SALAJU# = SARRASFE#
    CLSS$ = MKS$(CVS(CLSS$)) & (Mid$(CLSS$, 5, 2))
    'CLSS$ = CVS(CLSS$) & Mid$(CLSS$, 5, 2)
    VDEF$ = CLSS$ + MKD$(SARRASFE#) + MKD$(SALAJU#) + MKI$(0) + Space$(256) + MKI$(0)
    OBJE$ = "CARAJUCLIN"
    If MODOMONEI% = 2 Then
       OBJE$ = "CARAJUCLI-DOL"
       VDEF$ = VDEF$ + MKI$(1) + MKS$(1)
    End If
102 CLSS$ = OBJPLA$(OBJE$, VDEF$)
    If CLSS$ = "" Or CLSS$ = "<ESC>" Then GoTo 5
    SALAJU# = CVD(Mid$(CLSS$, 15, 8))
    ICONI% = CVI(Mid$(CLSS$, 281, 2)) ' MODIFIQUE (ANTES 279,2)
    If ICONI% < 1 Or ECOARCH$("CUECON", MKI$(ICONI%)) = "" Then
       Call MENSERR(24, "Falta Imputación Contable Ajuste")
       VDEF$ = CLSS$
       GoTo 102
    End If
    If ICONI% = CUECOCLI%(NCLI) Then
       Call MENSERR(24, "Cuenta Contable no Válida")
       VDEF$ = CLSS$
       GoTo 102
    End If
    '
    DIFESAL# = (Int(100 * (SALAJU# - SARRASFE#) + 0.5)) / 100
    DIFEDEU# = (Int(100 * (SALFINAL# - SALDEUTO# + DIFESAL#) + 0.5)) / 100
    '
    If Abs(DIFESAL#) < 0.005 Then
      If Abs(DIFEDEU#) < 0.005 Then
        Call MENSERR(24, "No hay Ajuste")
        GoTo 5
      End If
    End If
    '
    MONAJU% = 1 ' PESOS
    COTIMONEM# = 1
    If MODOMONEI% = 2 Then
      MONAJU% = CVI(Mid$(CLSS$, 281, 2))
      If MONAJU% < 1 Or MONAJU% > 6 Or ECOARCH$("TAMONED", MKI$(MONAJU%)) = "" Then
         Call MENSERR(24, "Moneda no Válida")
         VDEF$ = CLSS$
         GoTo 102
      End If
      '
      COTIMONEM# = CVS(Mid$(CLSS$, 283, 4))
      If COTIMONEM# < 0.01 Then
         Call MENSERR(24, "Falta Cotización")
         VDEF$ = CLSS$
         GoTo 102
      End If
      If MONAJU% = 1 Then
         If Abs(COTIMONEM# - 1) > 0.00001 Then
            Call MENSERR(24, "Cotización no Válida")
            VDEF$ = CLSS$
            GoTo 102
         End If
      End If
    End If
    '
    FVENI% = CVI(Mid$(CLSS$, 21, 2))
    FEX = FAJUI%
    FETX$ = FECHATEX$(FEX)
    NRECIB$ = "AS-" + Mid$(FETX$, 7, 2) + Mid$(FETX$, 4, 2) + Mid$(FETX$, 1, 2) + "-"
    '
'    SQLX$ = "SELECT * FROM CAJABANC "
'    SQLX$ = SQLX$ + "WHERE TipoMovim = 'AJUD' "
'    SQLX$ = SQLX$ + "AND CodiCom_Cor LIKE '" & NRECIB$ + "*' "
'    SQLX$ = SQLX$ + "ORDER BY CodiCom_Cor DESC"
    
    ORMA% = (-1)
    
    Dim RECITEMP As ADODB.Recordset
    SQLX$ = "SELECT max(CodiCom_Cor) as Maximo FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE TipoMovim = 'AJUD' "
    SQLX$ = SQLX$ + "AND CodiCom_Cor LIKE '" & NRECIB$ + "%' "
    Set RECITEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    If Not IsNull(RECITEMP!maximo) Then
        ORMA% = XVALO(Mid$(RECITEMP!maximo, 11, 2))
    End If
    RECITEMP.Close
    Set RECITEMP = Nothing

    'Set RECITEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    'With RECITEMP
      'On Error Resume Next
      '.MoveFirst
      'If Err < 1 Then
      '  ORMA% = XVALO(Mid$(!CodiCom_Cor, 11, 2))
      'End If
    'End With
    
    ORMX$ = TRIM$(Str$(ORMA% + 1))
    While Len(ORMX$) < 2: ORMX$ = "0" + ORMX$: Wend
    NRECIB$ = NRECIB$ + Left$(ORMX$, 2)
    '
    MONEMIA% = MONEMI%
    '
20  NC1% = NC%
    Call APERBASDAT("CABAN")
    '
    MONEMI% = MONAJU%
    CoComCor$ = NRECIB$
    CoComLar$ = NRECIB$
'    CodiEmp% = 0
'    CodiSuc% = SUC%
    NumeClie = NCLI
    NumeProv% = 0
    FechEmi = CVDAT(FAJUI%)
    '
    Referen$ = rasocli$(NCLI)
    FeContab = CVDAT(FAJUI%)
    FeAcredi = CVDAT(FAJUI%)
    NumeTal = 0
    NumeCom = XVALO(Mid$(NRECIB$, 4, 6))
    Observa$ = rasocli$(NC%)
    '
    CuentaCo% = CUECOINT%(CUEMADRE$(NCLI))
    TiMovi = "AJUD"            ' AJUSTE DE SALDO PROVEEDORES
    OpMovi = "AS"              ' AJUSTE DE SALDO
    VaMovi = "CCORR"           ' CUENTA CORRIENTE
    Descrip = "Aj.Saldo - " + rasocli$(NCLI)
    ImpoDeb# = DIFESAL#
    ValoDola# = ROUND#(DIFESAL# / COTIMONEM#)
    HAYAJU% = 0: If Abs(DIFESAL#) > 0.005 Then HAYAJU% = 1
    Call GRACABAN
    '
'4.- Graba la Contra-Partida en ICONI%
    CuentaCo% = ICONI%
    TiMovi = "AJUSALD"        ' AJUSTE DE SALDO PROVEEDORES
    OpMovi = "AS"              ' AJUSTE DE SALDO
    VaMovi = "NO-IMP"          ' NO IMPUTADO
    Descrip = "Aj.Saldo - " + rasocli$(NCLI)
    ImpoDeb# = (-1) * DIFESAL#
    Call GRACABAN
    '
'5.- Grabar Vencimientos del Ajuste ******************************************
    SIGNO% = 1
    IBruTot# = DIFESAL#
    IIvaTot# = 0
    '
    NuorVen% = 1
    FeVenc = CVDAT(FVENI%)
    ImVenc# = DIFEDEU#
    Call GRADEUPEN
    '
    'CajaBanc.Close
    '
    If HAYAJU% > 0 Then
      '
      MOTIVO$ = MOTIVO$ + Mid$(CLSS$, 25, 279)
      Call GRAREG("!OBSERVOF", Left$(MOTIVO$, 64), 1)
      Call GRAREG("!OBSERVOF", Mid$(MOTIVO$, 65, 64), 2)
      Call GRAREG("!OBSERVOF", Mid$(MOTIVO$, 129, 64), 3)
      Call GRAREG("!OBSERVOF", Mid$(MOTIVO$, 193, 64), 4)
      Call SETULTREG("!OBSERVOF", 4)
      Call CIERRARCH("!OBSERVOF")
      '
      FORIPRE$ = TRIM$(Control$("", "FORAJUSA"))
      If TRIM$(FORIPRE$) <> "" Then
         '
         FEX = HOY(HOS$)
         FECHDOC$ = FECHATEX$(FEX)
         NUMEDOC$ = NRECIB$ 'TRIM$(Str$(NUAJU&)) '
         TIPODOC$ = "Ajuste Saldo Cliente"
         'DESTIDOC% = NC%
         DESTIDOCNUE = NCLI
         '
         FEX = FAJUI%
         CODPARAM$(1) = "FECH-VEN": DOCPARAM$(1) = FECHATEX$(FEX)
         CODPARAM$(2) = "IMPO-PEN": DOCPARAM$(2) = CSTRING$(MKD$(SARRASFE#), 4, 12, 2, 2)
         CODPARAM$(3) = "IMPO-CAN": DOCPARAM$(3) = CSTRING$(MKD$(DIFESAL#), 4, 12, 2, 2)
         CODPARAM$(4) = "IMP-TOTA": DOCPARAM$(4) = CSTRING$(MKD$(SALAJU#), 4, 12, 2, 2)
         CAPARDOC% = 4
         CACOLTAB1% = 0
         CAITAB1% = 0
         '
         Call PRINTDOC("FORAJUSA")   ' Remito de Facturacion
         '
      End If
      Call SETULTREG("!OBSERVOF", 0)
      Call CIERRARCH("!OBSERVOF")
      '
    End If
    '
    MONEMI% = MONEMIA%
    Screen.MousePointer = 1
    Call COMUNI("Ajuste Completo")
    '
    GoTo 5
    '
99  Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    '
End Sub
'
Sub AJUSALDOL()
    '
10  Call SELECHO("DEUDOR1")
    NC% = XVALO(VALACT1$("DEUDOR1"))
    If NC% < 1 Then Exit Sub
    '
    Dim DOCUCA$(), SALPES#(), SALDOL#()
    CADOC% = 0
    RFF$ = RECFILT$("CARBIMON", 2, MKI$(NC%))
    For JKLY& = 1 To Len(RFF$) Step 4
      RECABI& = CVS(Mid$(RFF$, JKLY&, 4))
      YCB$ = REGLEIDO$("CARBIMON", RECABI&)
      If TRIM$(Mid$(YCB$, 13, 18)) <> "" Then
        If TRIM$(Mid$(YCB$, 13, 18)) = TRIM$(Mid$(YCB$, 45, 18)) Then
          CADOC% = CADOC% + 1
          ReDim Preserve DOCUCA$(CADOC%), SALPES#(CADOC%), SALDOL#(CADOC%)
          DOCUCA$(CADOC%) = TRIM$(Mid$(YCB$, 13, 18))
          SALPES#(CADOC%) = 0
          SALDOL#(CADOC%) = 0
        End If
      End If
    Next JKLY&
    '
    For JKLY& = 1 To Len(RFF$) Step 4
      RECABI& = CVS(Mid$(RFF$, JKLY&, 4))
      XXC$ = REGLEIDO$("CARBIMON", RECABI&)
      DOCABE$ = TRIM$(Mid$(XXC$, 13, 18))
      For KUI% = 1 To CADOC%
        If DOCUCA$(KUI%) = DOCABE$ Then
          SALPES#(KUI%) = SALPES#(KUI%) + (Int(100 * (CVD(Mid$(XXC$, 65, 8)) - CVD(Mid$(XXC$, 73, 8))) + 0.5)) / 100
          SALDOL#(KUI%) = SALDOL#(KUI%) + (Int(100 * (CVD(Mid$(XXC$, 97, 8)) - CVD(Mid$(XXC$, 105, 8))) + 0.5)) / 100
          GoTo 19
        End If
      Next KUI%
19  Next JKLY&
    '
    JJ& = 0
    For KUI% = 1 To CADOC%
      If Abs(SALPES#(KUI%)) < 0.005 Then
        If Abs(SALDOL#(KUI%)) >= 0.005 Then
          IMPOS$ = NUMSTRI$(SALDOL#(KUI%), 14, 2, 1)
          NPX& = XVALO(Mid$(DOCUCA$(KUI%), 11, 8))
          Y$ = MKS$(NPX&) + AJUSTI$(DOCUCA$(KUI%), 18) + IMPOS$
          JJ& = JJ& + 1
          Call GRAREG("!SELCUERE", Y$, JJ&)
        End If
      End If
    Next KUI%
    Call SETULTREG("!SELCUERE", JJ&)
    Call FRESHECHO("!SELCUERE")
    '
    If JJ& < 1 Then
      Call COMUNI("La Cuenta Elegida no Presenta\Comprobantes en Moneda Extranjera\que Puedan Ajustarse.")
      GoTo 10
    End If
    '
    Call SELECHO("!SELCUERE")
    REGCUEC& = XVALO(VALACT1$("!SELCUERE"))
    If XVALO(REGCUEC&) < 1 Then
        GoTo 10
    End If
    '
    DOCUCABE$ = TRIM$(Left$(VALACT2$("!SELCUERE"), 18))
    DIFEX# = XVALO(Mid$(VALACT2$("!SELCUERE"), 19, 14))
    If COMALTER%("Confirme el Ajuste a CERO del Saldo\en Moneda Extranjera del Documento '" + DOCUCABE$ + "'.\\Confirmar\Cancelar") = 1 Then
      RECILAR$ = "AJ-X-0000-00000000"
      RFF$ = RECFILT$("CARBIMON", 2, MKI$(NC%))
      For JKLY& = 1 To Len(RFF$) Step 4
        RECABI& = CVS(Mid$(RFF$, JKLY&, 4))
        YCB$ = REGLEIDO$("CARBIMON", RECABI&)
        If TRIM$(Mid$(YCB$, 13, 18)) = DOCUCABE$ Then
            MONEMORI% = CVI(Mid$(YCB$, 31, 2))
            TICAMBX# = TICAMONE(MONEMORI%)
            FEDOAPLI% = HOY(HOS$)
            NROO& = XVALO(Mid$(RECILAR$, 11))
          Call REPLA(YCB$, MKI$(FEDOAPLI%), 33, 2)
          Call REPLA(YCB$, MKI$(NC%), 35, 2)
          Call REPLA(YCB$, Chr$(TIDOAP%), 37, 1)
          Call REPLA(YCB$, Chr$(VADOAP%), 38, 1)
          Call REPLA(YCB$, MKS$(NROO&), 39, 4)
          Call REPLA(YCB$, MKI$(SUC%), 43, 2)
          Call REPLA(YCB$, RECILAR$, 45, 18)
          Call REPLA(YCB$, MKI$(MONEMORI%), 63, 2) ' RECIBO SIEMPRE EN PESOS
          Call REPLA(YCB$, MKD$(0), 65, 8)
          Call REPLA(YCB$, MKD$(0), 73, 8)
          Call REPLA(YCB$, MKD$(0), 81, 8)
          Call REPLA(YCB$, MKD$(TICAMBX#), 89, 8)
          Call REPLA(YCB$, String$(16, 0), 97, 16)
          If DIFEX# < 0 Then
              Call REPLA(YCB$, MKD$(Abs(DIFEX#)), 97, 8)
            Else
              Call REPLA(YCB$, MKD$(DIFEX#), 105, 8)
          End If
          Call REPLA(YCB$, MKD$(0), 113, 8)
          '
          Call REGAPP("CARBIMON", YCB$)
          Call CIERRARCH("CARBIMON")
          '
          Call RECUSALDOL(DOCUCABE$, NC%)
          '
          Call COMUNI("Ajuste Procesado")
          GoTo 10
        End If
      Next JKLY&
      '
      Call MENSERR(24, "No fue Posible Procesar el Ajuste")
    End If
    '
    GoTo 10
End Sub
'
Sub IMPUCREVEN(Optional NCLI)
    '
    If XVALO(NCLI) <= 0 Then
      NCLI = Abs(CLIPROACTIVO)
    End If
    If NCLI < 1 Then
'10      Call SELECHO("DEUDOR12")
'        NCLI = XVALO(VALACT1$("DEUDOR12"))
10      NCLI = ListaSeleccionClientes()
        If NCLI < 1 Then
            GoTo 999
        End If
    End If
    '
    FIMPUC07.Label1 = Str$(NCLI)
    FIMPUC07.Label2 = rasocli$(NCLI)
    FIMPUC07.Label3 = DOMICLI$(NCLI)
    FIMPUC07.Label6 = CPOSCLI$(NCLI)
    FIMPUC07.Label12 = LOCACLI$(NCLI)
    '
    FIMPUC07.List2.Clear
    Call BLANARCH("!SELDOCAP")
    Call FRESHECHO("!SELDOCAP")
    FIMPUC07.Refresh
    '
    
' AQUI ESCRIBIR UNA RUTINA PARA SELECCIONAR DE SALDACRE LOS DOCUMENTOS A IMPUTAR Y CARGARLOS EN ARCHIVO 'SELDOCAP'
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM SADEUPEN "
    SQLX$ = SQLX$ + "WHERE NuClien = " & NCLI & " "
    SQLX$ = SQLX$ + "ORDER BY FeVencim ASC;"
    'Set FacPenTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    'jandy sql
    Dim FacPenTemp As ADODB.Recordset
    Set FacPenTemp = New ADODB.Recordset
    FacPenTemp.CursorLocation = adUseClient
25  On Error Resume Next
    FacPenTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
    End If
    On Error GoTo 0
    '
    With FacPenTemp
      On Error Resume Next
      .MoveLast
      If Err < 1 Then
        .MoveFirst
        Do While Not .EOF
          TIX$ = UCase$(Left$(!codicom_lar, 2))
'          SG% = (-1)
'          TIPOS$ = "Factura": If TIX$ = "AS" Then TIPOS$ = "Ajuste de Saldo"
'                              If TIX$ = "TF" Then TIPOS$ = "Ticket-Factura"
'                              If TIX$ = "RH" Then TIPOS$ = "Rec.Honorarios"
'                              If TIX$ = "ND" Then TIPOS$ = "Nota de Debito"
'                              If TIX$ = "CO" Then TIPOS$ = "Comprobante"
'                              If TIX$ = "OP" Then TIPOS$ = "Orden de Pago": SG% = 1
'                              If TIX$ = "FF" Then TIPOS$ = "O/P Fondo Fijo": SG% = 1
'                              If TIX$ = "NC" Then TIPOS$ = "Nota de Credito": SG% = 1
'                              If TIX$ = "DP" Then TIPOS$ = "Debito Propio": SG% = 1
                              '
          IMPO# = !IBruComp
          'SUMAFAC# = SUMAFAC# + IMPO#
          '
          SALDO# = !SaldDebe - !SALDAcre
          SALDOL# = !SaldDebe_DOLA - !SaldAcre_DOLA
          If MODOMONEI% < 2 Then
             If Abs(SALDO#) < 0.005 Then
                SALDOL# = 0
             End If
          End If
          '
          MEMI% = !Mone_Emis
          'SUMADEU# = SUMADEU# + SALDO#
          '
          NROX$ = !codicom_lar
          FEMI = CVINT(!FECHEMISI)
          FEVE = CVINT(!FEVENCIM)
          '
          IMPTEX$ = CSTRING$(MKD$(IMPO#), 3, 12, 2, 2)
          SALTEX$ = CSTRING$(MKD$(Abs(SALDO#)), 3, 12, 2, 2)
          '
          If SALDO# >= 0.005 Or SALDOL# > 0 Then
               '
               SUMAFAC# = SUMAFAC# + IMPO#
               IMPO# = Abs(IMPO#)
               SUMADEU# = SUMADEU# + SALDO#
               '
               Z$ = Space$(128)
               Call REPLA(Z$, NROX$, 1, 18)
               Call REPLA(Z$, FECHATEX$(FEMI), 23, 8)
               Call REPLA(Z$, IMPTEX$, 32, 13)
               Call REPLA(Z$, FECHATEX$(FEVE), 47, 8)
               Call REPLA(Z$, SALTEX$, 57, 13)
               '        Call REPLA(Z$, REGTEX$, 119, 8)
               '        Call REPLA(Z$, IRETEX$, 127, 2)
               If MODOMONEI% = 2 Then GoTo 20
               '
               ' ESTO ES PARA PROCESO BIMONETARIO
               MONEMORI% = 0: SALDOLAX# = 0
               RFF$ = RECFILT$("CARBIMON", 2, MKI$(NCLI))
               DOCUCABE$ = TRIM$(NROX$)
               For JKLY& = 1 To Len(RFF$) Step 4
                 RECABI& = CVS(Mid$(RFF$, JKLY&, 4))
                 YCB$ = REGLEIDO$("CARBIMON", RECABI&)
                 If TRIM$(Mid$(YCB$, 13, 18)) = DOCUCABE$ Then
                   If TRIM$(Mid$(YCB$, 45, 18)) = DOCUCABE$ Then
                     MONEMORI% = CVI(Mid$(YCB$, 31, 2))
                   End If
                   SALDOLAX# = SALDOLAX# + CVD(Mid$(YCB$, 97, 8)) - CVD(Mid$(YCB$, 105, 8))
                   Call REPLA(YCB$, MKD$(0), 113, 8)
                 End If
               Next JKLY&
               If MONEMORI% > 1 Then
                 SALDO1# = SALDOLAX# * TICAMONE(MONEMORI%)
                 If SALDO1# < 0 Then SALDO1# = 0
                 If SALDO1# > SALDO# Then SALDO1 = SALDO#
                 SALTEX$ = CSTRING$(MKD$(SALDO1#), 3, 12, 2, 2)
                 Call REPLA(Z$, SALTEX$, 57, 13)
               End If
               ' FIN ESTO ES PARA PROCESO BIMONETARIO
               '
20             FIMPUC07.List2.AddItem Z$
               '
             ElseIf SALDO# <= (-0.005) Then
               '
               Z$ = Space$(60) + MKS$(0)
               Call REPLA(Z$, NROX$, 1, 18)
               Call REPLA(Z$, FECHATEX$(FEMI), 19, 8)
               Call REPLA(Z$, IMPTEX$, 27, 14)
               Call REPLA(Z$, SALTEX$, 41, 14)
               '     Call REPLA(Z$, MKS$(REG&), 61, 4)
               '     Call REPLA(Z$, MKI$(KKI%), 59, 2)
                    '
               Call REGAPP("!SELDOCAP", Z$)
               KQ& = KQ& + 1
               '
          End If
        .MoveNext
        Loop
      End If
    End With
    '
    If ULTREG&("!SELDOCAP") < 1 Then
      Call MENSERR(24, "No Hay Créditos Pendientes de Aplicación\en la Cuenta Elegida.")
      GoTo 10
    End If
    '
    SPDX$ = REGSEL$("!SELDOCAP")
    If Len(SPDX$) <= 4 Then GoTo 10
    '
    With FIMPUC07
      .Label27 = "Saldo no Imputado:"
      .Label16 = ""
      .Label16.Visible = False
      .Label31.Visible = False
      .Label30.Visible = False
      .Label28.Visible = False
      '
      NROXEL$ = TRIM$(Left$(SPDX$, 18))
      .Label25 = NROXEL$
      .Label24 = Mid$(SPDX$, 19, 8)
      .Label23 = TRIM$(Mid$(SPDX$, 27, 14))
      .Label22 = TRIM$(Mid$(SPDX$, 41, 14))
      .Label14 = TRIM$(Mid$(SPDX$, 41, 14))
      .Label18 = ""
      .Label20 = .Label14
      .Label10 = ""
      If MODOMONEI% = 2 Then
         NCX1% = XVALO(.Label1)
         FEMIX = FECHANUM(.Label24)
         MEMI% = MONEMISOR%(NROXEL$, FEMIX, NCX1%)
         TICAM# = TICAMISOR(NROXEL$, FEMIX, NCX1%)
         .Label23 = TRIM$(FORMATNUM$(ROUND#((XVALO(.Label23)) / TICAM#), "F12.2"))
         .Label22 = TRIM$(FORMATNUM$(ROUND#((XVALO(.Label22)) / TICAM#), "F12.2"))
         .Label14 = TRIM$(FORMATNUM$(ROUND#((XVALO(.Label14)) / TICAM#), "F12.2"))
         .Label27 = "Saldo N/I:"
         .Label16 = FORMATNUM$(TICAM#, "F9.4")
         .Label28 = SIMBOLPES$(MEMI%)
         .Label32 = TRIM$(Str$(MEMI%))
         .Label16.Visible = True
         .Label28.Visible = True
         .Label30.Visible = True
         .Label31.Visible = True
         '
         If MEMI% > 1 Then
            ' VOLVER A CARGAR EL LIST2
            .List2.Clear
            
            With FacPenTemp
               On Error Resume Next
               .MoveLast
               If Err < 1 Then
                  .MoveFirst
                  Do While Not .EOF
                     TIX$ = UCase$(Left$(!codicom_lar, 2))
                     IMPO# = !IBruComp
                     SALDO# = !SaldDebe_DOLA - !SaldAcre_DOLA
                     MONEDOCU% = !Mone_Emis
                     If MEMI% > 1 Then
                       If MEMI% <> MONEDOCU% Then GoTo 95
                     End If
                     NROX$ = !codicom_lar
                     FEMI = CVINT(!FECHEMISI)
                     FEVE = CVINT(!FEVENCIM)
                     IMPTEX$ = CSTRING$(MKD$(IMPO#), 3, 12, 2, 2)
                     SALTEX$ = CSTRING$(MKD$(Abs(SALDO#)), 3, 12, 2, 2)
                     '
                     If SALDO# >= 0.005 Then
                        Z$ = Space$(128)
                        Call REPLA(Z$, NROX$, 1, 18)
                        Call REPLA(Z$, FECHATEX$(FEMI), 23, 8)
                        Call REPLA(Z$, IMPTEX$, 32, 13)
                        Call REPLA(Z$, FECHATEX$(FEVE), 47, 8)
                        Call REPLA(Z$, SALTEX$, 57, 13)
                        FIMPUC07.List2.AddItem Z$
                     End If
95                .MoveNext
                  Loop
               End If
            End With
            ' FIN DE VOLVER A CARGAR EL LIST2
         End If
      End If
    End With
    FacPenTemp.Close
    Set FacPenTemp = Nothing
   
    '
    FIMPUC07.Show 1
    GoTo 10
    '
999 End Sub


Sub IMPUCREVEN_INCLUYE_AP_AUTOMATICA(Optional NCLI)
    '
    If XVALO(NCLI) <= 0 Then
      NCLI = Abs(CLIPROACTIVO)
    End If
    If NCLI < 1 Then
10      Call SELECHO("DEUDOR12")
        NCLI = XVALO(VALACT1$("DEUDOR12"))
        If NCLI < 1 Then
            GoTo 999
        End If
    End If
    '
    FIMPUC07.Label1 = Str$(NCLI)
    FIMPUC07.Label2 = rasocli$(NCLI)
    FIMPUC07.Label3 = DOMICLI$(NCLI)
    FIMPUC07.Label6 = CPOSCLI$(NCLI)
    FIMPUC07.Label12 = LOCACLI$(NCLI)
    '
    FIMPUC07.List2.Clear
    Call BLANARCH("!SELDOCAP")
    Call FRESHECHO("!SELDOCAP")
    FIMPUC07.Refresh
    ''********************************************************************************************************************************
    'ESTE BLOQUE ES PARA AGREGAR LOS RECIBOS A LA TABLA SADEUPEN QUE LOS MISMOS SE HAYAN BORRADO POR LA RECONSTRUCCION AUTOMATICA.
    '*********************************************************************************************************************************
'      RECORRO CAJABANC Y COMPARO CON APLICOBRA A MISMO CLIENTE Y RECIBO LA SUMA APLICADA CON EL VALOR ORIGINAL
'  SI ES MAYOR EL IMPORTE DEL RECIBO QUE LO (APLICADO + SADEUPEN)  THEN
'       SI EL RECIBO EXISTE EN SADEUPEN ACTUALIZO EL SALDACRE
'       SI EL RECIBO NO EXISTE REALIZO UN GRADEUPEN DEL RECIBO
'
    SQLX$ = " SELECT CODICOM_LAR, IHABECOMP,FECHEMISI  FROM CAJABANC WITH(NOLOCK) WHERE IHABECOMP > 0  AND NUCLIEN = " & NCLI
    SQLX$ = SQLX$ + " AND (TIPOMOVIM = 'FVEN' OR TIPOMOVIM = 'RCOB')"
    SQLX$ = SQLX$ + " ORDER BY FECHEMISI, CODICOM_LAR"
    Set RS = READSET(SQLX$)
    With RS
      Do While Not .EOF
        RECIBOX$ = SAFETEXT$(!codicom_lar)
        FE% = CVINT(!FECHEMISI)
        IMPORTEORIGINARL# = XVALO(!ihabecomp)
        YAPLICOBRA# = APLICADOENAPLICOBRA#(RECIBOX$, NCLI)
        ENSADEUP# = ENSADEUPEN#(RECIBOX$, NCLI)
        If (IMPORTEORIGINARL#) > (ENSADEUP# + YAPLICOBRA# + 0.05) Then
           If ENSADEUP# <= 0 Then 'CASO QUE NO ESTABA EN SADEUPEN LO AGREGA
                NROX1$ = TRIM$(Mid$(RECIBOX$, 13, 6))
                FEMIX1% = FE%
                'FC-002293-B
                CoComCor$ = Left$(RECIBOX$, 2) & "-" & NROX1$ & "-" & "00"
                CoComLar$ = RECIBOX$
                NumeClie = NCLI
                NumeProv% = 0
                FechEmi = CVDAT(FEMIX1%)
                Referen$ = rasocli$(NCLI)
                NumeCom = XVALO(NROX1$)
                Observa$ = rasocli$(NCLI)
                '
                IBruTot# = (-1) * IMPORTEORIGINARL#
                IIvaTot# = 0
                NuorVen% = 0
                FeVenc = CVDAT(FEMIX1%)
                ImVenc# = (-1) * IMPORTEORIGINARL#
                Call GRADEUPEN
           Else
                Condi$ = "CODICOM_LAR = '" & RECIBOX$ & "'" And NUCLIEN = " & NCLI"
                Call ACTUREGISTRO("SADEUPEN", "SALDACRE", Condi$, FORMATNUM$(IMPORTEORIGINARL# - YAPLICOBRA#, "F16.5"), REGAF&, "NOMESS")
           End If
        End If
        .MoveNext
      Loop
    End With
    '/////////////////////////////////////////////////
    ''********************************************************************************************************************************
    'ESTE BLOQUE ES PARA AGREGAR LAS FACTURAS A LA TABLA SADEUPEN QUE LOS MISMOS SE HAYAN BORRADO POR LA RECONSTRUCCION AUTOMATICA.
    '*********************************************************************************************************************************
    
    'RECORRO EL CAJABANC LAS FACTURAS  Y COMPARO CON APLICOBRA A MISMO CLIENTE Y FACTURA (CODOCOR_LAR)
'SI EL IMPORTE DE LA FACTURA ES > LA SUMA DE LO APLICADO EN APLICOBRA + SADEUPEN
'  SI EXISTE EN SADEUPEN ACTUALIZO EL SALDACRE
'  SI NO EXISTE REALIZO UN GRADEUFA.

    SQLX$ = " SELECT CODICOM_LAR, IDEBECOMP,FECHEMISI, CondPag  FROM CAJABANC WITH(NOLOCK) WHERE IDEBECOMP > 0  AND NUCLIEN = " & NCLI
    SQLX$ = SQLX$ + " AND (TIPOMOVIM = 'FVEN' OR TIPOMOVIM = 'AJ') "
    SQLX$ = SQLX$ + " ORDER BY FECHEMISI, CODICOM_LAR"
    Set RS = READSET(SQLX$)
    With RS
      Do While Not .EOF
        FACTUX$ = SAFETEXT$(!codicom_lar)
        FE% = CVINT(!FECHEMISI)
        IMPORTEORIGINARL# = XVALO(!idebecomp)
        YAPLICOBRA# = FACT_APLICADOENAPLICOBRA#(FACTUX$, NCLI)
        CONPAX% = XVALO(!CondPag)
        ENSADEUP# = FAC_ENSADEUPEN#(FACTUX$, NCLI)
        If (IMPORTEORIGINARL#) > (ENSADEUP# + YAPLICOBRA# + 0.05) Then
           If ENSADEUP# <= 0 Then 'CASO QUE NO ESTABA EN SADEUPEN LO AGREGA
                NROX1$ = TRIM$(Mid$(FACTUX$, 13, 6))
                FEMIX1% = FE%
                'FC-002293-B
                TIPOX% = 2
                If Left$(FACTUX$, 2) = "ND" Then TIPOX% = 3
                If Left$(FACTUX$, 2) = "NC" Then TIPOX% = 6
                letx$ = Right(FACTUX$, 1)
                VARI% = 1
                If letx$ = "B" Then VARI% = 2
                If letx$ = "E" Then VARI% = 3
                NROX1$ = TRIM$(Mid$(FACTUX$, 11, 6))
                'FC-002293-B
                CoComCor$ = Left$(FACTUX$, 2) & "-" & NROX1$ & "-" & letx$ 'SIN PREFIJO
                CoComLar$ = FACTUX$
                NumeClie = NCLI
                NumeProv% = 0
                FechEmi = CVDAT(FEMIX1%)
                Referen$ = rasocli$(NCLI)
                NumeCom = XVALO(NROX1$)
                Observa$ = rasocli$(NCLI)
                '
                IBruTot# = IMPORTEORIGINARL#
                If CONPAX% <= 0 Then CONPAX% = CPAGCLI%(NCLI)
                Call FORFASQL.GRADEUFA(TIPOX%, VARI%, NROX1$, NCLI, FEMIX1%, CONPAX%, IMPORTEORIGINARL# - YAPLICOBRA#)
                If IMPORTEORIGINARL# <> (IMPORTEORIGINARL# - YAPLICOBRA#) Then
                   Condi$ = "CODICOM_LAR = '" & FACTUX$ & "' And NUCLIEN = " & NCLI
                   Call ACTUREGISTRO("SADEUPEN", "IBRUCOMP", Condi$, FORMATNUM$(IMPORTEORIGINARL#, "F16.5"), REGAF&, "NOMESS")
                End If
           Else
                Condi$ = "CODICOM_LAR = '" & FACTUX$ & "'" And NUCLIEN = " & NCLI"
                Call ACTUREGISTRO("SADEUPEN", "SALDDEBE", Condi$, FORMATNUM$(IMPORTEORIGINARL# - YAPLICOBRA#, "F16.5"), REGAF&, "NOMESS")
           End If
        End If
        .MoveNext
      Loop
    End With
'////////////////////////////////////////////////////////////////////
    ''********************************************************************************************************************************
    'FIN DE DESCOMPOSICION AUTOMATICA.
    '*********************************************************************************************************************************


' AQUI ESCRIBIR UNA RUTINA PARA SELECCIONAR DE SALDACRE LOS DOCUMENTOS A IMPUTAR Y CARGARLOS EN ARCHIVO 'SELDOCAP'
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM SADEUPEN "
    SQLX$ = SQLX$ + "WHERE NuClien = " & NCLI & " "
    SQLX$ = SQLX$ + "ORDER BY FeVencim ASC;"
    'Set FacPenTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    'jandy sql
    Dim FacPenTemp As ADODB.Recordset
    Set FacPenTemp = New ADODB.Recordset
    FacPenTemp.CursorLocation = adUseClient
25  On Error Resume Next
    FacPenTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
    End If
    On Error GoTo 0
    '
    With FacPenTemp
      On Error Resume Next
      .MoveLast
      If Err < 1 Then
        .MoveFirst
        Do While Not .EOF
          TIX$ = UCase$(Left$(!codicom_lar, 2))
'          SG% = (-1)
'          TIPOS$ = "Factura": If TIX$ = "AS" Then TIPOS$ = "Ajuste de Saldo"
'                              If TIX$ = "TF" Then TIPOS$ = "Ticket-Factura"
'                              If TIX$ = "RH" Then TIPOS$ = "Rec.Honorarios"
'                              If TIX$ = "ND" Then TIPOS$ = "Nota de Debito"
'                              If TIX$ = "CO" Then TIPOS$ = "Comprobante"
'                              If TIX$ = "OP" Then TIPOS$ = "Orden de Pago": SG% = 1
'                              If TIX$ = "FF" Then TIPOS$ = "O/P Fondo Fijo": SG% = 1
'                              If TIX$ = "NC" Then TIPOS$ = "Nota de Credito": SG% = 1
'                              If TIX$ = "DP" Then TIPOS$ = "Debito Propio": SG% = 1
                              '
          IMPO# = !IBruComp
          'SUMAFAC# = SUMAFAC# + IMPO#
          '
          SALDO# = !SaldDebe - !SALDAcre
          SALDOL# = !SaldDebe_DOLA - !SaldAcre_DOLA
          If MODOMONEI% < 2 Then
             If Abs(SALDO#) < 0.005 Then
                SALDOL# = 0
             End If
          End If
          '
          MEMI% = !Mone_Emis
          'SUMADEU# = SUMADEU# + SALDO#
          '
          NROX$ = SAFETEXT$(!codicom_lar) ' RECIBO
          FEMI = CVINT(!FECHEMISI)
          FEVE = CVINT(!FEVENCIM)
          '
          IMPTEX$ = CSTRING$(MKD$(IMPO#), 3, 12, 2, 2)
          SALTEX$ = CSTRING$(MKD$(Abs(SALDO#)), 3, 12, 2, 2)
          '
          If SALDO# >= 0.005 Or SALDOL# > 0 Then
               '
               SUMAFAC# = SUMAFAC# + IMPO#
               IMPO# = Abs(IMPO#)
               SUMADEU# = SUMADEU# + SALDO#
               '
               Z$ = Space$(128)
               Call REPLA(Z$, NROX$, 1, 18)
               Call REPLA(Z$, FECHATEX$(FEMI), 23, 8)
               Call REPLA(Z$, IMPTEX$, 32, 13)
               Call REPLA(Z$, FECHATEX$(FEVE), 47, 8)
               Call REPLA(Z$, SALTEX$, 57, 13)
               '        Call REPLA(Z$, REGTEX$, 119, 8)
               '        Call REPLA(Z$, IRETEX$, 127, 2)
               If MODOMONEI% = 2 Then GoTo 20
               '
               ' ESTO ES PARA PROCESO BIMONETARIO
               MONEMORI% = 0: SALDOLAX# = 0
               RFF$ = RECFILT$("CARBIMON", 2, MKI$(NCLI))
               DOCUCABE$ = TRIM$(NROX$)
               For JKLY& = 1 To Len(RFF$) Step 4
                 RECABI& = CVS(Mid$(RFF$, JKLY&, 4))
                 YCB$ = REGLEIDO$("CARBIMON", RECABI&)
                 If TRIM$(Mid$(YCB$, 13, 18)) = DOCUCABE$ Then
                   If TRIM$(Mid$(YCB$, 45, 18)) = DOCUCABE$ Then
                     MONEMORI% = CVI(Mid$(YCB$, 31, 2))
                   End If
                   SALDOLAX# = SALDOLAX# + CVD(Mid$(YCB$, 97, 8)) - CVD(Mid$(YCB$, 105, 8))
                   Call REPLA(YCB$, MKD$(0), 113, 8)
                 End If
               Next JKLY&
               If MONEMORI% > 1 Then
                 SALDO1# = SALDOLAX# * TICAMONE(MONEMORI%)
                 If SALDO1# < 0 Then SALDO1# = 0
                 If SALDO1# > SALDO# Then SALDO1 = SALDO#
                 SALTEX$ = CSTRING$(MKD$(SALDO1#), 3, 12, 2, 2)
                 Call REPLA(Z$, SALTEX$, 57, 13)
               End If
               ' FIN ESTO ES PARA PROCESO BIMONETARIO

20             FIMPUC07.List2.AddItem Z$
               '
             ElseIf SALDO# <= (-0.005) Then
               '
               Z$ = Space$(60) + MKS$(0)
               Call REPLA(Z$, NROX$, 1, 18)
               Call REPLA(Z$, FECHATEX$(FEMI), 19, 8)
               Call REPLA(Z$, IMPTEX$, 27, 14)
               Call REPLA(Z$, SALTEX$, 41, 14)
               '     Call REPLA(Z$, MKS$(REG&), 61, 4)
               '     Call REPLA(Z$, MKI$(KKI%), 59, 2)
                    '
               Call REGAPP("!SELDOCAP", Z$)
               KQ& = KQ& + 1
               '
          End If
        .MoveNext
        Loop
      End If
    End With
    '
    
    If ULTREG&("!SELDOCAP") < 1 Then
      Call MENSERR(24, "No Hay Créditos Pendientes de Aplicación\en la Cuenta Elegida.")
      GoTo 10
    End If
    '
    Call SETULTREG("!SELDOCAP", KQ&)
    SPDX$ = REGSEL$("!SELDOCAP")
    If Len(SPDX$) <= 4 Then GoTo 10
    '
    With FIMPUC07
      .Label27 = "Saldo no Imputado:"
      .Label16 = ""
      .Label16.Visible = False
      .Label31.Visible = False
      .Label30.Visible = False
      .Label28.Visible = False
      '
      NROXEL$ = TRIM$(Left$(SPDX$, 18))
      .Label25 = NROXEL$
      .Label24 = Mid$(SPDX$, 19, 8)
      .Label23 = Mid$(SPDX$, 27, 14)
      .Label22 = Mid$(SPDX$, 41, 14)
      .Label14 = Mid$(SPDX$, 41, 14)
      .Label18 = ""
      .Label20 = .Label14
      .Label10 = ""
      If MODOMONEI% = 2 Then
         NCX1% = XVALO(.Label1)
         FEMIX = FECHANUM(.Label24)
         MEMI% = MONEMISOR%(NROXEL$, FEMIX, NCX1%)
         TICAM# = TICAMISOR(NROXEL$, FEMIX, NCX1%)
         .Label23 = FORMATNUM$(ROUND#((XVALO(.Label23)) / TICAM#), "F12.2")
         .Label22 = FORMATNUM$(ROUND#((XVALO(.Label22)) / TICAM#), "F12.2")
         .Label14 = FORMATNUM$(ROUND#((XVALO(.Label14)) / TICAM#), "F12.2")
         .Label27 = "Saldo N/I:"
         .Label16 = FORMATNUM$(TICAM#, "F9.4")
         .Label28 = SIMBOLPES$(MEMI%)
         .Label32 = TRIM$(Str$(MEMI%))
         .Label16.Visible = True
         .Label28.Visible = True
         .Label30.Visible = True
         .Label31.Visible = True
         '
         If MEMI% > 1 Then
            ' VOLVER A CARGAR EL LIST2
            .List2.Clear
            
            With FacPenTemp
               On Error Resume Next
               .MoveLast
               If Err < 1 Then
                  .MoveFirst
                  Do While Not .EOF
                     TIX$ = UCase$(Left$(!codicom_lar, 2))
                     IMPO# = !IBruComp
                     SALDO# = !SaldDebe_DOLA - !SaldAcre_DOLA
                     MONEDOCU% = !Mone_Emis
                     If MEMI% > 1 Then
                       If MEMI% <> MONEDOCU% Then GoTo 95
                     End If
                     NROX$ = !codicom_lar
                     FEMI = CVINT(!FECHEMISI)
                     FEVE = CVINT(!FEVENCIM)
                     IMPTEX$ = CSTRING$(MKD$(IMPO#), 3, 12, 2, 2)
                     SALTEX$ = CSTRING$(MKD$(Abs(SALDO#)), 3, 12, 2, 2)
                     '
                     If SALDO# >= 0.005 Then
                        Z$ = Space$(128)
                        Call REPLA(Z$, NROX$, 1, 18)
                        Call REPLA(Z$, FECHATEX$(FEMI), 23, 8)
                        Call REPLA(Z$, IMPTEX$, 32, 13)
                        Call REPLA(Z$, FECHATEX$(FEVE), 47, 8)
                        Call REPLA(Z$, SALTEX$, 57, 13)
                        FIMPUC07.List2.AddItem Z$
                     End If
                     '
95                .MoveNext
                  Loop
               End If
            End With
            ' FIN DE VOLVER A CARGAR EL LIST2
         End If
      End If
    End With
    FacPenTemp.Close
    Set FacPenTemp = Nothing
    '
    '
    FIMPUC07.Show 1
    GoTo 10
    '
999 End Sub

Sub DESAPLIREC()
    '
    'jandy sql
    
10  Call SELECHO("DEUDOR1")
    NCLI = XVALO(VALACT1$("DEUDOR1"))
    If NCLI < 1 Then
        GoTo 999
    End If
    '
    Call BLANARCH("!SELCUERE")
    Call FRESHECHO("!SELCUERE")
    '
' AQUI ESCRIBIR UNA RUTINA PARA SELECCIONAR DE SALDACRE LOS DOCUMENTOS A IMPUTAR Y CARGARLOS EN ARCHIVO 'SELDOCAP'
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT DISTINCT CoDocAp_Lar, FechEmisi FROM ApliCobra "
    SQLX$ = SQLX$ + "WHERE NuClien = " & NCLI & " "
    SQLX$ = SQLX$ + "ORDER BY FechEmisi ASC;"
    
    'jandy sql se podria mejorar con un group by
    Dim FacPenTemp As ADODB.Recordset
    Set FacPenTemp = FLEXCONN.Execute(FILTSQL$(SQLX$))
    
    'Set FacPenTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    JJ& = 0
    On Error Resume Next
    With FacPenTemp
     If Not (.EOF And .BOF) Then
        .MoveFirst
        Do While Not .EOF
        TIX$ = !codocap_lar
        SUAPLI# = 0
        
        'SQLX$ = "NuClien = " & NCLI & " AND CoDocAp_Lar = '" & TIX$ & "'"
         '
        strSQL = " Select sum(ImpoApli) as Suapli from Aplicobra where CoDocAp_Lar = '" & TIX$ & "'"
        Dim Aplicobra As ADODB.Recordset
        Set Aplicobra = FLEXCONN.Execute(FILTSQL$(strSQL))
        SUAPLI# = Aplicobra!SUAPLI
        
'          Aplicobra.FindFirst SQLX$
'20        If Aplicobra.NoMatch = False Then
'
'            SUAPLI# = SUAPLI# + Aplicobra!ImpoApli
'            Aplicobra.FindNext SQLX$
'            GoTo 20
'          End If
'          '
          IMPOS$ = NUMSTRI$(SUAPLI#, 14, 2, 1)
          NPX& = XVALO(Mid$(!codocap_lar, 11, 8))
          Y$ = MKS$(NPX&) + AJUSTI$(!codocap_lar, 18) + IMPOS$
          JJ& = JJ& + 1
          Call GRAREG("!SELCUERE", Y$, JJ&)
          '
        .MoveNext
        Loop
      End If
    End With
    FacPenTemp.Close
    Set FacPenTemp = Nothing
    Aplicobra.Close
    Set Aplicobra = Nothing
    '
    
    If ULTREG&("!SELCUERE") < 1 Then
      Call MENSERR(24, "No Hay Créditos Aplicados\en la Cuenta Elegida.")
      GoTo 10
    End If
    '
    SPDX$ = REGSEL$("!SELCUERE")
    If Len(SPDX$) <= 4 Then GoTo 10
    '
    TIX$ = Mid$(SPDX$, 5, 18)
    'SQLX$ = "NuClien = " & NCLI & " AND CoDocAp_Lar = '" & TIX$ & "'"
    '
    Call DesaplicaReciboEnAnotadorPedidos(NCLI, TIX$)
    
    strSQL = "Delete ApliCobra where "
    strSQL = strSQL & "NuClien = " & NCLI & " AND CoDocAp_Lar = '" & TIX$ & "'"
    Call EJECUTACOMANDO(strSQL)
    '
    '========================PARA FACTURA ANTICIPO====================
    'ESTO ES EN EL CASO QUE LA NOTA DE CREDITO ESTE ASOCIADA A UNA FACTURA ANTICIPO LA ELIMINE.
    If Left$(TRIM$(TIX$), 3) = "NC-" Then
       strSQL = "Delete FACTUANTI where "
       strSQL = strSQL & " NuClien = " & NCLI & " AND DOCU_APLI = '" & TIX$ & "'"
       Call EJECUTACOMANDO(strSQL)
    End If
    Condi$ = " NuClien = " & NCLI & " AND TIPOMOVIM = 'AjReci'  and left(descrimov ,18) =  '" & TIX$ & "' "
    Call BORRAREGISTROS("CAJABANC", Condi$, REGAF&)
    
    '========================PARA FACTURA ANTICIPO====================
    
    
'    With Aplicobra
'30    .FindFirst SQLX$
'      If .NoMatch = False Then
'        .Delete
'        GoTo 30
'      End If
'    End With
    '
    
    DOCULA$ = TIX$
    RFG$ = RECFILT$("CARBIMON", 2, MKI$(NCLI))
    For KUL& = Len(RFG$) - 3 To 1 Step -4
      RECABIN& = CVS(Mid$(RFG$, KUL&, 4))
      XXC$ = REGLEIDO$("CARBIMON", RECABIN&)
      If TRIM$(Mid$(XXC$, 45, 18)) = TRIM$(DOCULA$) Then
        Call GRAREG("CARBIMON", String$(256, 0), RECABIN&)
      End If
    Next KUL&
    '
    Call COMUNI("Des-Aplicación de " + TRIM$(TIX$) + " Completa.\Debe Ahora Reconstruir la Composición de Deuda del Cliente.")
    '
    RECUEC07.Text1.TEXT = TRIM$(Str$(NCLI))
    Hide
    Call FRESHALL
    PROLLAMA$ = Name
    Call RECUEC07.Command4_Click
    RECUEC07.Show 1
    Call FRESHALL
    '
    GoTo 10
    '
999 End Sub
'
Sub REPIRECIBO()
10  Screen.MousePointer = 1
'    Call SELECHO("DEUDOR1")
'    NCLI = XVALO(VALACT1$("DEUDOR1"))
    NCLI = ListaSeleccionClientes()
    If NCLI < 1 Then
        GoTo 999
    End If
    '
    Screen.MousePointer = 11
    FENUE$ = Control$("", "FENURECI")
    FLIXX% = FECHANUM(FENUE$)
    If FLIXX% > HOY(HO$) Then FLIXX% = HOY(HO$)
    Call BLOQARCH("!SELCUERE")
    Call BLANARCH("!SELCUERE")
    '
301 Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + " WHERE TipoMovim= 'RCOB' "
    SQLX$ = SQLX$ + " AND NuClien= " & NCLI & " "
    'Set COBRANTMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
   Dim COBRANTMP As ADODB.Recordset
   Set COBRANTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With COBRANTMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        On Error GoTo 0
        Call MENSERR(24, "Sin Movimientos")
        GoTo 10
      End If
      On Error GoTo 0
      Y$ = REGBLAN("!SELCUERE")
      JJ& = 0
      Do While Not .EOF
        NROLL& = !NumeComp
        IMPO# = !ValAbsComp
        FC% = CVINT(!FECHEMISI)
        '
        If NROLL& > 0 Then
          IMPOS$ = NUMSTRI$(IMPO#, 13, 2, 1)
          NNXX# = CDbl(NROLL&)
          DOCU$ = "Recibo No. " + TRIM$(Str$(NROLL&))
          Call REPLA(Y$, MKS$(!NumeComp), 1, 4)
          Call REPLA(Y$, DOCU$, 5, 18)
          Call REPLA(Y$, IMPOS$, 23, 13)
          JJ& = JJ& + 1
          Call GRAREG("!SELCUERE", Y$, JJ&)
        End If
        .MoveNext
      Loop
    End With
    COBRANTMP.Close
    Set COBRANTMP = Nothing
    
    Call CIERRARCH("!SELCUERE")
    Call FRESHECHO("!SELCUERE")
    Screen.MousePointer = 1
    '
    If JJ& < 1 Then
      Call MENSERR(24, "Sin Recibos Posteriores al " + FENUE$)
      GoTo 10
    End If
    '
350 Call SELECHO("!SELCUERE")
    REGCUEC& = XVALO(VALACT1$("!SELCUERE"))
    If XVALO(REGCUEC&) < 1 Then
        GoTo 999
    End If
    '
    NROLL& = 0
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE NumeComp = " & REGCUEC & " "
    SQLX$ = SQLX$ + " AND TipoMovim= 'RCOB'"
    'Set COBRANTMPB = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim COBRANTMPB As ADODB.Recordset
    Set COBRANTMPB = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With COBRANTMPB
        On Error Resume Next
        .MoveFirst
        If Err Then
          On Error GoTo 0
          Call MENSERR(24, "Imposible Abrir Documento Desde esta Aplicación")
          GoTo 10
        End If
        On Error GoTo 0
        NCLII = !NUCLIEN
        NROLL& = !NumeComp
    End With
    COBRANTMPB.Close
    Set COBRANTMPB = Nothing
    '
    If NROLL& = 0 Then
        Call COMUNI("Recibo Nro. " & CStr(NROLL&) & " No Encontrado")
        GoTo 350
    End If
    '
    TIDOCUM$ = "Recibo Nro."
    NUDOCU$ = TRIM$(Str$(NROLL&))
    DOCUNU& = 0
'    RACLI$ = TRIM$(Left$(RASOCLI$(NCLI%), 8))
'    '
'    Screen.MousePointer = 11
'    For UI& = ULTREG&("PDOCLST") To 1 Step -1
'      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
'      If InStr(XX$, TIDOCUM$) > 4 Then
'        If InStr(XX$, NUDOCU$) > 4 Then
'          If InStr(XX$, RACLI$) > 16 Then
'            DOCUNU& = CVS(Left$(XX$, 4))
'            GoTo 5
'          End If
'        End If
'      End If
'    Next UI&
'    Screen.MousePointer = 1
'    Call COMUNI("Imposible Abrir Documento Desde esta Aplicación")
'    GoTo 10
'    '
5   Call MUESTRADOC(DOCUNU&, , TIDOCUM$, NUDOCU$, NUDOCU$)
    GoTo 350
    '
999 Call CIERRARCH("*.*")
    Call LIBARCH("*.*")
    Screen.MousePointer = 1
    '
End Sub
'\\\OT-05-0441-WAR-30/06/05///
Sub VAPLIREC(Optional CLIACTIVO%)
    '
10  Screen.MousePointer = 1
    '\\\OT-07-0189-OD-24/04/08///
    If CLIACTIVO% > 0 Then NCLI = CLIACTIVO%: GoTo 20
    '\\\OT-07-0189-OD-FIN///
'    Call SELECHO("DEUDOR1")
'    NCLI = XVALO(VALACT1$("DEUDOR1"))
    NCLI = ListaSeleccionClientes()
20  If NCLI < 1 Then
        GoTo 999
    End If
    '
    Screen.MousePointer = 11
    Call BLOQARCH("!SELCUERE")
    Call BLANARCH("!SELCUERE")
    '
        '
    Call APERBASDAT("CABAN")
    SQLX$ = " SELECT  * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE NuClien = " & NCLI & " "
    SQLX$ = SQLX$ + "AND ( TipoMovim = 'RCOB' OR (TipoMovim = 'FVEN'AND OpciMovim = 'NC' ) ) "
    SQLX$ = SQLX$ + "ORDER BY CodiCom_Lar ASC "
    
    'jandy sql
    
    'Set TOTMP = CueCorri.OpenRecordset(SQLX$, 4)
    Dim TOTMP As ADODB.Recordset
    Set TOTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
'    Call APERBASDAT("CABAN")
'    SQLX$ = " SELECT  * FROM APLICOBRA "
'    SQLX$ = SQLX$ + " WHERE NuClien = " & NCLI & ""
'    SQLX$ = SQLX$ + " ORDER BY CoDocAp_Lar ASC "
'    Set APCOTMP = CueCorri.OpenRecordset(SQLX$, 4)
    '
'    With TOTMP
'      On Error Resume Next
'      .MoveFirst 'habia un movenext
'      If Err Then
'        On Error GoTo 0
'        Call MENSERR(24, "Sin Movimientos")
'        GoTo 10
'      End If
'      On Error GoTo 0
    With TOTMP
    If (.EOF And .BOF) Then
      Call MENSERR(24, "Sin Movimientos")
      If CLIACTIVO% > 0 Then Exit Sub 'SI CLIENTE SELECCIONADO PREVIO ,SALE
      GoTo 10
    End If
      Y$ = REGBLAN("!SELCUERE")
      JJ& = 0
      CODVIE$ = ""
      
      Do While Not .EOF
        NROX$ = !codicom_lar   ' !CoDocAp_Lar
        IMPO# = !ihabecomp     ' 0
        Call REPLA(Y$, NROX$, 5, 18)
        If IMPO# >= 0.005 Then
          IMPOS$ = NUMSTRI$(IMPO#, 13, 2, 1)
          Call REPLA(Y$, IMPOS$, 23, 13)
          JJ& = JJ& + 1
          Call REPLA(Y$, MKS(JJ&), 1, 4)
          Call GRAREG("!SELCUERE", Y$, JJ&)
        End If
      .MoveNext
      Loop
      Call SETULTREG("!SELCUERE", JJ&)
    End With
    TOTMP.Close
    Set TOTMP = Nothing
    
    '
    Call SETULTREG("!SELCUERE", JJ&)
    Call CIERRARCH("!SELCUERE")
    Call FRESHECHO("!SELCUERE")
    Screen.MousePointer = 1
    '
310 RESEL$ = REGSEL$("!SELCUERE/" + TRIM$(rasocli$(NCLI)))
    If Len(RESEL$) > 4 Then
      DAPLI$ = Mid$(RESEL$, 5, 18)
      RESTO# = XVALO(Mid$(RESEL$, 23, 13))
      ENCAP$ = TRIM$(DAPLI$) + " - $ " + TRIM$(Mid$(RESEL$, 23, 13)) + " - " + rasocli$(NCLI)
      '
      SQLX$ = " SELECT * FROM APLICOBRA "
      SQLX$ = SQLX$ + " WHERE NuClien = " & NCLI & ""
      SQLX$ = SQLX$ + " AND  CoDocAp_Lar= '" & TRIM(DAPLI$) & "'"
      SQLX$ = SQLX$ + " ORDER BY CoDocOr_Lar ASC "
      Dim APCOTMP As ADODB.Recordset
      Set APCOTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
      'Set APCOTMP = CueCorri.OpenRecordset(SQLX$, 4)
      '
      JJ& = 0
      With APCOTMP    ' ApliCobra
      '  .FindFirst "CoDocAp_Lar= '" & TRIM(DAPLI$) & "'"
307   '  If .NoMatch = False Then
      If Not (.EOF And .BOF) Then
        .MoveFirst
        Do While Not (.EOF)
          Y$ = REGBLAN$("!SELDOCAP")
          NROY$ = !codocor_lar
          REFY$ = Left(!refeapli, 18)
          IMPO# = !IMPOAPLI
            If IMPO# > RESTO# Then IMPO# = RESTO#
          If IMPO# >= 0.005 Then
            IMPOS$ = NUMSTRI$(IMPO#, 13, 2, 1)
            Call REPLA(Y$, NROY$, 1, 18)
            Call REPLA(Y$, REFY$, 19, 24)
            Call REPLA(Y$, IMPOS$, 44, 13)
            JJ& = JJ& + 1
            Call GRAREG("!SELDOCAP", Y$, JJ&)
            RESTO# = RESTO# - IMPO#
          End If
          '.FindNext "CoDocAp_Lar= '" & TRIM(DAPLI$) & "'"
          'GoTo 307
          .MoveNext
        Loop
      End If
      End With
      APCOTMP.Close
      Set APCOTMP = Nothing
      '
      If RESTO# >= 0.005 Then
        Y$ = REGBLAN$("!SELDOCAP")
        NROY$ = "Pendiente de Aplicación"
        IMPOS$ = NUMSTRI$(RESTO#, 13, 2, 1)
        Call REPLA(Y$, NROY$, 19, 24)
        Call REPLA(Y$, IMPOS$, 44, 13)
        JJ& = JJ& + 1
        Call GRAREG("!SELDOCAP", Y$, JJ&)
      End If
      '
      Call SETULTREG("!SELDOCAP", JJ&)
      Call FRESHECHO("!SELDOCAP")
      '
      RESEL$ = REGSEL$("!SELDOCAP/" + ENCAP$)
      GoTo 310
      '
    End If
    If CLIACTIVO% > 0 Then Exit Sub 'SI CLIENTE SELECCIONADO PREVIO ,SALE
    GoTo 10
    '
999 End Sub

Sub VAPLIFAC()
      '
10  Screen.MousePointer = 1
'    Call SELECHO("DEUDOR1")
'    NCLI = XVALO(VALACT1$("DEUDOR1"))
    NCLI = ListaSeleccionClientes()
    If NCLI < 1 Then
        GoTo 999
    End If
    '
    Screen.MousePointer = 11
    Call BLOQARCH("!SELCUERE")
    Call BLANARCH("!SELCUERE")
    '
    Call APERBASDAT("CABAN")
    SQLX$ = " SELECT  * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE NuClien = " & NCLI & " "
    SQLX$ = SQLX$ + "AND TipoMovim = 'FVEN' AND OpciMovim <> 'NC'"
    SQLX$ = SQLX$ + "ORDER BY CodiCom_Lar ASC "
    'Set TOTMP = CueCorri.OpenRecordset(SQLX$, 4)
    'jandy sql
    Dim TOTMP As ADODB.Recordset
    Set TOTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))

    
    '
'    Call APERBASDAT("CABAN")
'    SQLX$ = " SELECT  * FROM APLICOBRA "
'    SQLX$ = SQLX$ + " WHERE NuClien = " & NCLI & ""
'    SQLX$ = SQLX$ + " ORDER BY CoDocOr_LAR ASC "
'    Set APCOTMP = CueCorri.OpenRecordset(SQLX$, 4)
    '
'    With TOTMP
'      On Error Resume Next
'      .MoveFirst
'      If Err Then
'        On Error GoTo 0
'        Call MENSERR(24, "Sin Movimientos")
'        GoTo 10
'      End If
'      On Error GoTo 0
    With TOTMP
      If (.EOF And .BOF) Then
        Call MENSERR(24, "Sin Movimientos")
        GoTo 10
      End If
      Y$ = REGBLAN("!SELCUERE")
      JJ& = 0
      Do While Not .EOF
        NROX$ = !codicom_lar
        IMPO# = TOTMP!idebecomp
        '
        If IMPO# >= 0.005 Then
          Call REPLA(Y$, NROX$, 5, 18)
            IMPOS$ = NUMSTRI$(IMPO#, 13, 2, 1)
          Call REPLA(Y$, IMPOS$, 23, 13)
          JJ& = JJ& + 1
          Call REPLA(Y$, MKS(JJ&), 1, 4)
          Call GRAREG("!SELCUERE", Y$, JJ&)
        End If
      .MoveNext
      Loop
      Call SETULTREG("!SELCUERE", JJ&)
    End With
    TOTMP.Close
    Set TOTMP = Nothing
    
    '
    Call SETULTREG("!SELCUERE", JJ&)
    Call CIERRARCH("!SELCUERE")
    Call FRESHECHO("!SELCUERE")
    Screen.MousePointer = 1
    '
310 RESEL$ = REGSEL$("!SELCUERE/" + TRIM$(rasocli$(NCLI)))
    If Len(RESEL$) > 4 Then
      DAPLI$ = Mid$(RESEL$, 5, 18)
      RESTO# = XVALO(Mid$(RESEL$, 23, 13))
      ENCAP$ = TRIM$(DAPLI$) + " - $ " + TRIM$(Mid$(RESEL$, 23, 13)) + " - " + rasocli$(NCLI)
      '
      SQLX$ = " SELECT  * FROM APLICOBRA "
      SQLX$ = SQLX$ + " WHERE NuClien = " & NCLI & ""
      SQLX$ = SQLX$ + " AND CoDocOR_LAR= '" & TRIM(DAPLI$) & "'"
      SQLX$ = SQLX$ + " ORDER BY CoDocOr_LAR ASC "
      Dim APCOTMP As ADODB.Recordset
      Set APCOTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
      
      JJ& = 0
      With APCOTMP
        If Not (.EOF And .BOF) Then
          .MoveFirst
        Do While Not (.EOF)
'        .FindFirst "CoDocOR_LAR= '" & TRIM(DAPLI$) & "'"
'307     If .NoMatch = False Then
          Y$ = REGBLAN$("!SELDOCAP")
          NROY$ = !codocap_lar
          REFY$ = Left(!refeapli, 18)
          IMPO# = !IMPOAPLI
            If IMPO# > RESTO# Then IMPO# = RESTO#
          If IMPO# >= 0.005 Then
            IMPOS$ = NUMSTRI$(IMPO#, 13, 2, 1)
            Call REPLA(Y$, NROY$, 1, 18)
            Call REPLA(Y$, REFY$, 19, 24)
            Call REPLA(Y$, IMPOS$, 44, 13)
            JJ& = JJ& + 1
            Call GRAREG("!SELDOCAP", Y$, JJ&)
            RESTO# = RESTO# - IMPO#
          End If
          '.FindNext "CoDocOr_Lar= '" & TRIM(DAPLI$) & "'"
          'GoTo 307
          .MoveNext
          Loop
        End If
      End With
      APCOTMP.Close
      Set APCOTMP = Nothing
      
      '
      If RESTO# >= 0.005 Then
        Y$ = REGBLAN$("!SELDOCAP")
        NROY$ = "Pendiente de Aplicación"
        IMPOS$ = NUMSTRI$(RESTO#, 13, 2, 1)
        Call REPLA(Y$, NROY$, 19, 24)
        Call REPLA(Y$, IMPOS$, 44, 13)
        JJ& = JJ& + 1
        Call GRAREG("!SELDOCAP", Y$, JJ&)
      End If
      '
      Call SETULTREG("!SELDOCAP", JJ&)
      Call FRESHECHO("!SELDOCAP")
      '
      RESEL$ = REGSEL$("!SELDOCAP/" + ENCAP$)
      GoTo 310
      '
    End If
    GoTo 10
    '
999 End Sub
'
Sub CONAPLI()
  '
  Call APERBASDAT("CABAN")

  SQLX$ = " SELECT * FROM APLICOBRA "
  'Set DOCAPTMP = CueCorri.OpenRecordset(SQLX$, 4)
  Dim DOCAPTMP As ADODB.Recordset
  Set DOCAPTMP = New ADODB.Recordset
20 On Error Resume Next
  DOCAPTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
  If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 20
  End If
  On Error GoTo 0
  '
  With DOCAPTMP
    '
    On Error Resume Next
    .MoveLast
    If Err Then
      On Error GoTo 0
      Exit Sub
    End If
    On Error GoTo 0
    Max& = .RecordCount
    JJ& = 0
    YY$ = REGBLAN("APLICOB")
    FEC1% = 0
    .MoveFirst
    Do While Not .EOF
      Call REPLA(YY$, MKI(!NUCLIEN), 3, 2)
      Call REPLA(YY$, !codocap_lar, 5, 18)
      Call REPLA(YY$, MKD(!IMPOAPLI), 23, 8)
      Call REPLA(YY$, Left(!refeapli, 24), 31, 24)
      Call REPLA(YY$, !codocor_lar, 55, 18)
       DOC1$ = !codocap_lar
       sqlx1$ = " SELECT FechEmisi, CodiCom_Lar FROM CAJABANC "
       sqlx1$ = sqlx1$ + " WHERE CodiCom_Lar = '" & DOC1$ & "'"
      'Set NROTMP = CueCorri.OpenRecordset(sqlx1$, 4)
      Dim NROTMP As ADODB.Recordset
      Set NROTMP = New ADODB.Recordset
23    On Error Resume Next
      NROTMP.Open FILTSQL$(sqlx1$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
      If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 23
      End If
      On Error GoTo 0
      '
      With NROTMP
         On Error Resume Next
         '.FindFirst "CodiCom_Lar = '" & DOC1$ & "' "
         'If .NoMatch = True Then
         If (.EOF And .BOF) Then
          On Error GoTo 0
          GoTo 25
         End If
         On Error GoTo 0
          FEC1% = CVINT(!FECHEMISI)
       End With
25    Call REPLA(YY$, MKI(FEC1%), 1, 2)
      JJ& = JJ& + 1
      Call TRACE(24, JJ&, Max&)
      Call GRAREG("APLICOB", YY$, JJ&)
      .MoveNext
    Loop
    '
  End With
  DOCAPTMP.Close
  Set DOCAPTMP = Nothing
  NROTMP.Close
  Set NROTMP = Nothing
  
  Call SETULTREG("APLICOB", JJ&)
  Call CIERRARCH("*.*")
  '
End Sub
'\\\OT-05-0441-WAR-FIN///
'\\\OT-05-0441-WAR-FIN///
Sub LIAPLIFE()
'jandy sql
   SOLONOTACREDITO% = 0
   OPX% = COMALTER%("Opciones de Consulta:\\Todos\Solo Notas de Crédito")
   If OPX% < 1 Or OPX% > 2 Then Exit Sub
   If OPX% = 2 Then SOLONOTACREDITO% = 1
   '
   Call DESHASFECHA(DES%, HAS%, PERIO$)
   If PERIO$ = "" Then
      GoTo 99
   End If
   '
   Screen.MousePointer = 11
   '
   'DESA = Int(CDbl(CVDAT(DES%)))
   'HASA = Int(CDbl(CVDAT(HAS%)))
   
   DESA = FETEXCOR1$(DES%)
   HASA = FETEXCOR1$(HAS%)
   '
   Call APERBASDAT("CABAN")
   SQLX$ = " SELECT NUCLIEN, IHabeComp, FechEmisi, CodiCom_Lar FROM CAJABANC WITH(NOLOCK) "
   'SQLX$ = SQLX$ + " WHERE INT(CDBL(FechEmisi)) >= " & DESA & " "
   'SQLX$ = SQLX$ + " AND INT(CDBL(FechEmisi)) <= " & HASA & " "
   SQLX$ = SQLX$ + " WHERE FechEmisi >= '" & DESA & "' "
   SQLX$ = SQLX$ + " AND FechEmisi <= '" & HASA & "' "
   If SOLONOTACREDITO% = 1 Then
        SQLX$ = SQLX$ + " AND (TipoMovim = 'FVEN'AND OpciMovim = 'NC' ) "
   Else
        SQLX$ = SQLX$ + " AND ( TipoMovim = 'RCOB' OR (TipoMovim = 'FVEN'AND OpciMovim = 'NC' ) ) "
   End If
   SQLX$ = SQLX$ + " Order by FechEmisi,CodiCom_Lar ASC "
   'Set NROTMP = CueCorri.OpenRecordset(SQLX$, 4)
   '
   'SQLX$ = " SELECT * FROM APLICOBRA "
   'Set DOCAPTMP = CueCorri.OpenRecordset(SQLX$, 4)
   '
   Dim NROTMP As ADODB.Recordset
   Set NROTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))

   With NROTMP
     On Error Resume Next
     .MoveFirst
     If Err Then GoTo 90
     On Error GoTo 0
     '
     Do While .EOF = False
       FEDOC% = CVINT(!FECHEMISI)
       NCLI = !NUCLIEN
       DOCAPLI$ = !codicom_lar
       IMPOTO# = !ihabecomp
       RESTO# = IMPOTO#
       If IMPOTO# < 0.005 Then GoTo 26
       '
       PRIREN% = 1
       
       'SQLX2$ = " CODOCAP_LAR = '" & DOCAPLI$ & "' "
       'DOCAPTMP.FindFirst SQLX2$

       strSQL = "Select * from Aplicobra WITH(NOLOCK) where CODOCAP_LAR = '" & DOCAPLI$ & "' "
       Dim DOCAPTMP As ADODB.Recordset
       Set DOCAPTMP = FLEXCONN.Execute(FILTSQL$(strSQL))
       With DOCAPTMP
       If Not (.EOF And .BOF) Then
        .MoveFirst
        Do While Not .EOF
'36     If DOCAPTMP.NoMatch = False Then
         YY$ = REGBLAN("APLICO1")
         Call REPLA(YY$, MKI(FEDOC%), 1, 2)
         Call REPLA(YY$, DOCAPLI$, 3, 18)
         Call REPLA(YY$, MKS(NCLI), 21, 4)
         If PRIREN% = 1 Then
           Call REPLA(YY$, MKD(IMPOTO#), 25, 8)
           PRIREN% = 0
         End If
         Call REPLA(YY$, DOCAPTMP!codocor_lar, 33, 18)
           REFEA$ = Left(DOCAPTMP!refeapli, 24)
         Call REPLA(YY$, REFEA$, 51, 24)
           IAPLI# = DOCAPTMP!IMPOAPLI
           If IAPLI# > RESTO# Then IAPLI# = RESTO#
         Call REPLA(YY$, MKD(IAPLI#), 75, 8)
         RESTO# = RESTO# - IAPLI#
         If IAPLI# >= 0.005 Then
           JJ& = JJ& + 1
           Call GRAREG("APLICO1", YY$, JJ&)
           CAREGRA% = CAREGRA% + 1
         End If
'         DOCAPTMP.FindNext SQLX2$
'         GoTo 36
         .MoveNext
         Loop
       End If
       End With
       DOCAPTMP.Close
       Set DOCAPTMP = Nothing
       '
       ' ESTO ES PARA PREVER QUE ESE RECIBO AUN NO SE APLICó NADA
       If RESTO# >= 0.005 Then
         
         ' buscar si aun está pendiente de aplicacion
         REFEB$ = "*** Pend.Aplic.***"
         REFEA$ = "Pendiente Aplic."
         
         'SaDeuPen.FindFirst "Nuclien = " & NCLI & " AND CodiCom_Lar = '" & DOCAPLI$ & "' AND (SaldDebe + SaldAcre) > 0.005"
         strSQL = "Select * from SaDeuPen WITH(NOLOCK) where Nuclien = " & NCLI
         strSQL = strSQL & " AND CodiCom_Lar = '" & DOCAPLI$ & "'"
         strSQL = strSQL & " AND (SaldDebe + SaldAcre) > 0.005"
         Dim SaDeuPen As ADODB.Recordset
         Set SaDeuPen = FLEXCONN.Execute(FILTSQL$(strSQL))
         With SaDeuPen
            If (.EOF And .BOF) Then
                ' ESTO ES PARA CONTEMPLAR LOS CASOS DONDE AL COINCIDIR LOS IMPORTES DE UNA NC Y UNA FACTURA, LA APLICACION
                ' NO SE GUARDA EN LA TABLA APLICOBRA. POR EJEMPLO, CUANDO SE UTILIZA LA OPCION 'ANULA FACTURA'
                strSQL = "Select CODICOM_LAR from CAJABANC WITH(NOLOCK) where ABS(IDEBECOMP - " & TRIM$(FORMATNUM$(IMPOTO#, "F10.2")) & ")< 0.005"
                strSQL = strSQL & " AND FECHCONTA <='" & FETEXCOR1$(FEDOC%) & "' "
                strSQL = strSQL & " AND NUCLIEN ='" & CStr(NCLI) & "' AND (TipoMovim = 'FVEN' AND OpciMovim = 'FC')"
                Dim RST As ADODB.Recordset
                Set RST = FLEXCONN.Execute(FILTSQL$(strSQL))
                With RST
                    If Not (.EOF And .BOF) Then
                        .MoveFirst
                         RESTO# = IMPOTO#
                         REFEB$ = SAFETEXT$(RST!codicom_lar)
                         REFEA$ = "Aplicación Auto."
                    End If
                End With
                RST.Close
                Set RST = Nothing
            End If
         End With
         SaDeuPen.Close
         Set SaDeuPen = Nothing
         
         YY$ = REGBLAN("APLICO1")
         Call REPLA(YY$, MKI(FEDOC%), 1, 2)
         Call REPLA(YY$, DOCAPLI$, 3, 18)
         Call REPLA(YY$, MKS(NCLI), 21, 4)
         If PRIREN% = 1 Then
            Call REPLA(YY$, MKD(IMPOTO#), 25, 8)
         End If
         Call REPLA(YY$, REFEB$, 33, 18)
         Call REPLA(YY$, REFEA$, 51, 24)
         Call REPLA(YY$, MKD(RESTO#), 75, 8)
         JJ& = JJ& + 1
         Call GRAREG("APLICO1", YY$, JJ&)
       End If
       '
26   .MoveNext
     Loop
     '
   End With
   NROTMP.Close
   Set NROTMP = Nothing
   
   
   '
90 Call SETULTREG("APLICO1", JJ&)
   Call HEADERS("APLICO1", "")
   Call HEADERS("APLICO1", "Corresponde a: " + TRIM$(PERIO$))
   Call FINAL("*LIAPLIC1")
   '
99 Screen.MousePointer = 1
   '
End Sub

Private Sub Form_Load()
   '
   UTILIZA_SKIN% = 1
   Call APLICACIONSKINS(Me, Picture1)
   MODOMONEI% = 0
   MODOMONE$ = Left$(TRIM$(UCase$(Control$("GESVENTA", "MODOMONE"))), 8)
   If MODOMONE$ = "MULTIMON" Then
      MODOMONEI% = 2
   End If
   '
End Sub
Sub CONTROLAPLI()

    If DERACCE%("CONIMPU", "Control de Imputaciones") < 2 Then GoTo 99

    Call APERBASDAT("CABAN")
    '
    FIN& = CantRegistros&("APLICOBRA")
    U& = 0
    '
    Screen.MousePointer = 11
    SQLX$ = " SELECT * FROM APLICOBRA WITH(NOLOCK)"
    Dim CABATMP As ADODB.Recordset
    Set CABATMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With CABATMP
      Do While Not .EOF
        U& = U& + 1
        Call TRACE(20, U&, FIN&)
        NUMECLIXX% = XVALO(!NUCLIEN)
        NUMEFACTU$ = SAFETEXT(!codocor_lar)
        SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK)"
        SQLX$ = SQLX$ + " WHERE codicom_lar = '" & NUMEFACTU$ & "'"
        Dim RSTMP As ADODB.Recordset
        Set RSTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
        With RSTMP
            Do While Not .EOF ' si el numero de factura no esta en el
               'cajabanc no entra por que el proceso esta verificando
               'solo las que existen en los 2
               If NUMECLIXX% = XVALO(!NUCLIEN) Then ' si coincide el numero de cliente sale de do while si no sigue recorriendo
                 Exit Do
               Else
                .MoveNext
                If .EOF = True Then 'si es fin de archivo es que no encontro coincidencia y presenta mensaje
                 Screen.MousePointer = 1
                 Call MENSERR(24, "No Coincide Cliente: " & NUMECLIXX% & " En Factura: " & NUMEFACTU$)
                 Screen.MousePointer = 11
                End If
               End If
              Loop
        End With
        RSTMP.Close
        Set RSTMP = Nothing
        .MoveNext
      Loop
    End With
    '
    CABATMP.Close
    Set CABATMP = Nothing
99  Screen.MousePointer = 1
    Call MENSERR(24, "Proceso Terminado")
    '
End Sub

Sub RECALCULO_DE_DIAS_DE_COBRANZA()
        ' CALCULO DE DIAS DE COBRANZA A PARTIR DE LA FECHA DE VENCIMIENTO
        ' DE LAS FACTURAS CANCELADAS
        ' CALCULO DE DIAS DE COBRANZA
        '
        ' COMPLETO LOS LIST CON LOS DATOS NECESARIOS PARA RECALCULAR LOS DIAS DE COBRANZA
        If CLACONTRA% < 1 Then                      ' iNGRESO DE CLAVE Y CONTRASEÑA
            Call MENSERR(24, "Santo Y Seña Inválido")
            Exit Sub
        End If
        '
        Call SELMESA(DES%, HAS%, PERIO$)
        If PERIO$ = "" Then GoTo 99
        DESDE$ = FETEXCOR1$(DES%)
        HASTA$ = FETEXCOR1$(HAS%)
        '
        Screen.MousePointer = 11
        SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK) WHERE FECHCONTA>='" & DESDE$ & "' AND FECHCONTA<='" & HASTA$ & "'"
        SQLX$ = SQLX$ & " AND TIPOMOVIM='RCOB'"
        Set RST = READSET(SQLX$)
        With RST
            Do While Not .EOF
                List1.Clear
                List2.Clear
                TOTAVALO# = 0
                RECI$ = SAFETEXT$(!codicom_lar)
                NC% = XVALO(!NUCLIEN)
                FECHA% = CVINT(!fechconta)
                sqlx1$ = "SELECT * FROM CAJABANC WITH(NOLOCK) WHERE CODICOM_LAR='" & RECI$ & "' AND TIPOMOVIM='CHRECI'"
                Set rst1 = READSET(sqlx1$)
                With rst1
                Do While Not .EOF
                    Z$ = Space$(96)
                    Call REPLA(Z$, CSTRING$(MKI$(1), 1, 4, 0, 2), 1, 4)
                    Call REPLA(Z$, BANCO$, 7, 14)
                    Call REPLA(Z$, NUVA$, 23, 12)
                    Call REPLA(Z$, PROTER$, 38, 2)
                    Call REPLA(Z$, CSTRING$(MKI$(CLEA%), 1, 4, 0, 2), 42, 4)
                        FEMI = CVINT(!fechconta)
                    Call REPLA(Z$, FECHATEX$(FEMI), 50, 8)
                        FEVE = CVINT(!fechacred)
                    Call REPLA(Z$, FECHATEX$(FEVE), 62, 8)
                        IMPTEX$ = FORMATNUM$(XVALO(!ValAbsComp), "F14.2")
                    Call REPLA(Z$, IMPTEX$, 70, 14)
                    TOTAVALO# = TOTAVALO# + XVALO(IMPTEX$)
                    List1.AddItem Z$
                    .MoveNext
                Loop
                End With
                
                sqlx2$ = "SELECT * FROM APLICOBRA WITH(NOLOCK) WHERE CODOCAP_LAR='" & RECI$ & "'"
                Set rst2 = READSET(sqlx2$)
                With rst2
                Do While Not .EOF
                    Z$ = Space$(128)
                        NROX$ = SAFETEXT$(!codocor_lar)
                    Call REPLA(Z$, NROX$, 1, 20)
                        Condi$ = "CODICOM_LAR='" & NROX$ & "' AND TIPOMOVIM='FVEN'"
                        FEMI = CVINT(VALOBADA("CAJABANC", "FECHCONTA", Condi$))
                    Call REPLA(Z$, FECHATEX$(FEMI), 23, 8)
                    Call REPLA(Z$, IMPTEX$, 32, 13)
                        NC% = XVALO(!NUCLIEN)
                        CONDICION_DE_PAGO% = CPAGCLI%(NC%)
                        CAD% = RECIBO08.PROMEDIO_CONDI_PAGOS(CONDICION_DE_PAGO%) 'DEVUELVE EL PROMEDIO DE LAS CONDICIONES DE PAGO
                        FEVE = DIASPOST(FEMI, CAD%)
                    Call REPLA(Z$, FECHATEX$(FEVE), 47, 8)
                        SALTEX$ = FORMATNUM(XVALO(!IMPOAPLI), "F14.2")
                    Call REPLA(Z$, SALTEX$, 56, 14)
                        ACRTEX$ = SALTEX$
                    Call REPLA(Z$, ACRTEX$, 70, 14)
                    Call REPLA(Z$, REGTEX$, 119, 8)
                    Call REPLA(Z$, IRETEX$, 127, 2)
                    List2.AddItem Z$
                    .MoveNext
                Loop
                End With
                '
                JJ% = 0: ACRED# = 0: TOTALVAL# = 0
                FechaRecibo% = FECHA%
                For II% = 0 To List1.ListCount - 1
                     ' VERIFICA QUE TIPO DE MEDIO ES. PARA ESO LEE EL CODIGO DE MEDIO
                     ' (LA ULTIMA COLUMNA DE LA TABLA DE MEDIOS DE COBRANZA)
                     ' OBVIA A LAS TRANSFERENCIAS (TR)
                     'MPAGO% = XVALO(Mid$(List1.List(II%), 1, 4))
                     'RGMP$ = FILTERGETRECORD$("MECOBRA", 1, MKI$(MPAGO%))
                     'CODMP$ = TRIM$(Mid$(RGMP$, 57, 4))
                     '
                     FVENCHE% = FECHANUM(Mid$(List1.List(II%), 62, 8))   ' fecha de vencimiento del cheque
                     'If FVENCHE% < FECHARECIBO% And CODMP$ <> "TR" Then FVENCHE% = FECHARECIBO%
                     IMPOCHE# = XVALO(Mid$(List1.List(II%), 70, 14))   ' importe del cheque
                     
5                    If ACRED# < 0.005 Then     ' SI HAY IMPORTE DE FACTURA PENDIENTE
10                       If JJ% < List2.ListCount Then ' RECORRE LOS COMPROBANTES DE LA COMPOSICION DE DEUDA
                             REGFACTURA$ = List2.List(JJ%)
                             FVENFAC% = FECHANUM(Mid$(REGFACTURA$, 23, 8)) ' FECHA DE EMISION DE LA FACTURA
                             ACRED# = XVALO(Mid$(REGFACTURA$, 70, 14))
                             If ACRED# < 0.005 Then    ' VERIFICA SI EL COMPROBANTE FUE ELEGIDO
                                 JJ% = JJ% + 1
                                 GoTo 10
                             End If
                         Else ' SI NO HAY MAS COMPROBANTES AGREGA EL SALDO DEL ULTIMO CHEQUE CON PONDERACION 1 (UNO)
                             DIFEDIA% = 1
                             TOTALVAL# = TOTALVAL# + (DIFEDIA% * IMPOCHE#)
                             Exit For
                         End If
                     End If
                         
                     If IMPOCHE# >= ACRED# Then ' EL IMPORTE DEL CHEQUE ES MAYOR AL DE LA FACTURA
                         IMPOCHE# = IMPOCHE# - ACRED#
                         '
                         DIFEDIA% = DIENTRE(FVENFAC%, FVENCHE%)
                         If ACRED# < 0 Then DIFEDIA% = 1 ' SI EL COMPROBANTE ES UN CREDITO, DIFEDIA% = 1
                         '
                         TOTALVAL# = TOTALVAL# + (DIFEDIA% * ACRED#)
                         ACRED# = 0: JJ% = JJ% + 1
                     Else                       ' EL IMPORTE DE LA FACTURA ES MAYOR AL DEL CHEQUE
                         DIFEDIA% = DIENTRE(FVENFAC%, FVENCHE%)
                         TOTALVAL# = TOTALVAL# + (DIFEDIA% * IMPOCHE#)
                         ACRED# = ACRED# - IMPOCHE#
                         IMPOCHE# = 0
                     End If
                     
                     If IMPOCHE# > 0.005 Then GoTo 5 ' SI TODAVIA HAY IMPORTE DEL CHEQUE, BUSCA UN NUEVO COMPROBANTE
                Next II%
                '
                'Label35 = FORMATNUM(0, "I1")
                'TOTAVALO# = XVALO(Label14) ' TOTAL DE VALORES RECIBIDOS
                If TOTAVALO# > 0.005 Then
                    AA1$ = FORMATNUM$((TOTALVAL# / TOTAVALO#), "I10") ' CALCULO DE DIAS PONDERADO
                    Call ACTUREGISTRO("CAJABANC", "ATRACHEQUE", "CODICOM_LAR='" & RECI$ & "' AND TIPOMOVIM='RCOB'", AA1$, REGAF&)
                End If
                '
                .MoveNext
            Loop
    End With
    
    Call COMUNI("Proceso de Recalculo de Dias de Cobranza Finalizado")
99  Screen.MousePointer = 1

End Sub

