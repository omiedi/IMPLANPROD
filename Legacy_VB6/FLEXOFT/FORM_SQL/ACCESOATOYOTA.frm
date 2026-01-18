VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form ACCESOATOYOTA 
   BackColor       =   &H00D0F0F0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acceso a Toyota"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   330
      Index           =   0
      Left            =   3240
      ScaleHeight     =   270
      ScaleWidth      =   1110
      TabIndex        =   26
      Top             =   480
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture8 
      BackColor       =   &H00006060&
      Height          =   7155
      Left            =   4515
      ScaleHeight     =   7095
      ScaleWidth      =   4215
      TabIndex        =   16
      Top             =   -420
      Width           =   4275
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   20
         Top             =   3735
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   21
            Top             =   0
            Width           =   12000
         End
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
         Picture         =   "ACCESOATOYOTA.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   700
         Width           =   650
      End
      Begin VB.CommandButton Command7 
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
         Picture         =   "ACCESOATOYOTA.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   1390
         Width           =   650
      End
      Begin VB.CommandButton Command1 
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
         Picture         =   "ACCESOATOYOTA.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Configuracion de Funcionamiento"
         Top             =   2415
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "ACCESOATOYOTA.frx":0896
         Top             =   3990
         Width           =   1515
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00D0F0F0&
      Caption         =   "Cartas de Embarque"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3840
      Index           =   2
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   4110
      Begin VB.PictureBox Picture9 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1300
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   1
         Top             =   3045
         Width           =   2430
         Begin VB.CommandButton Command14 
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
            Left            =   0
            Picture         =   "ACCESOATOYOTA.frx":27B8
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Consulta Detalle por Numero de  Remito"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Consulta Detalle por Nro. Remito"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   750
            Left            =   735
            TabIndex        =   3
            Top             =   105
            Width           =   1605
         End
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1100
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   4
         Top             =   2520
         Width           =   2430
         Begin VB.CommandButton Command8 
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
            Left            =   0
            Picture         =   "ACCESOATOYOTA.frx":2902
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Consulta / Listado de Detalle por Carta de Embarque"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Consulta por Car- ta de Embarque"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   750
            Left            =   735
            TabIndex        =   6
            Top             =   60
            Width           =   1605
         End
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   900
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   7
         Top             =   1995
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
            Height          =   600
            Left            =   0
            Picture         =   "ACCESOATOYOTA.frx":2C0C
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Genera Remito Automático de Carta de Embarque"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Generar Remito Automático"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   750
            Left            =   735
            TabIndex        =   9
            Top             =   60
            Width           =   1755
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   700
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   23
         Top             =   1470
         Width           =   2430
         Begin VB.CommandButton Command6 
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
            Left            =   0
            Picture         =   "ACCESOATOYOTA.frx":2D56
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Imprime Vale de Preparacion"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Imprimir Vale de Preparación"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   750
            Left            =   840
            TabIndex        =   25
            Top             =   60
            Width           =   1605
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   500
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   10
         Top             =   945
         Width           =   2430
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
            Height          =   600
            Left            =   0
            Picture         =   "ACCESOATOYOTA.frx":3060
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Guardar Carta de Embarque en BD FLEXOFT o Planilla Excel"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Guardar Carta de Embarque"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   750
            Left            =   840
            TabIndex        =   12
            Top             =   60
            Width           =   1380
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   300
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   13
         Top             =   420
         Width           =   2430
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
            Height          =   600
            Left            =   0
            Picture         =   "ACCESOATOYOTA.frx":336A
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Acceso a Portal OnLineToyota.com.ar"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Acceso al Por- tal TOYOTA"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   840
            TabIndex        =   15
            Top             =   60
            Width           =   1515
         End
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "ACCESOATOYOTA.frx":58F0
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "ACCESOATOYOTA.frx":5B24
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label1 
      Height          =   330
      Left            =   315
      TabIndex        =   22
      Top             =   3885
      Width           =   1905
   End
End
Attribute VB_Name = "ACCESOATOYOTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Command1.Enabled = False
   OPX% = ALTERNA%("Configura Acceso a Portal Toyota\Configura Acceso a EXCEL\Configuración Vale de Preparación\Otras Opciones de Configuración")
   If OPX% = 1 Then
        FORCONSHK.Label7 = ""
        FORCONSHK.Label6 = "0"
        ADIR$ = "APORTOYO"
        FORCONSHK.Text1 = VALCONTROL(0, ADIR$, "DESTAPLI")
        FORCONSHK.Text2 = VALCONTROL(0, ADIR$, "LINCOMMA")
        FORCONSHK.Text3 = VALCONTROL(0, ADIR$, "ARCHICON")
        FORCONSHK.Text4 = VALCONTROL(0, ADIR$, "TOOTIPTE")
        '
        FORCONSHK.Show 1
        If FORCONSHK.Label7 = "OK" Then
           DESTAPLI$ = FORCONSHK.Text1
           LINCOMMA$ = FORCONSHK.Text2
           Call GRAVALCON(0, ADIR$, "DESTAPLI", DESTAPLI$)
           Call GRAVALCON(0, ADIR$, "LINCOMMA", LINCOMMA$)
        End If
     ElseIf OPX% = 2 Then
        FORCONSHK.Label7 = ""
        FORCONSHK.Label6 = "0"
        ADIR$ = "ACCEXCEL"
        FORCONSHK.Text1 = VALCONTROL(0, ADIR$, "DESTAPLI")
        FORCONSHK.Text2 = VALCONTROL(0, ADIR$, "LINCOMMA")
        FORCONSHK.Text3 = VALCONTROL(0, ADIR$, "ARCHICON")
        FORCONSHK.Text4 = VALCONTROL(0, ADIR$, "TOOTIPTE")
        '
        FORCONSHK.Show 1
        If FORCONSHK.Label7 = "OK" Then
           DESTAPLI$ = FORCONSHK.Text1
           Call GRAVALCON(0, ADIR$, "DESTAPLI", DESTAPLI$)
        End If
     ElseIf OPX% = 3 Then
        CONFORVAL.Show 1
   End If
   '
   Command1.Enabled = True
End Sub

Private Sub Command2_Click()
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   OPX% = COMALTER%("Opciones de Guardar:\\En la Base de Datos FLEXOFT\En Planilla EXCEL")
   If OPX% = 1 Then
       Call GUARDACARTA
     Else
       AREXE$ = VALCONTROL$(0, "ACCEXCEL", "DESTAPLI")
       On Error Resume Next
       NX1 = Shell(AREXE$, vbNormalFocus)
       If Err Then
          MsgBox "Imposible Ejecutar '" + AREXE$ + " " + LINCOMMA$ + "'"
       End If
       On Error GoTo 0
   End If
   Command3.Enabled = True
End Sub

Sub GUARDACARTA()
   With CARTEMBAR07.GRID
     .Cols = 8
     .Rows = 1
     .Row = 0
     .COL = 0
     .ColWidth(0) = 600
     .ColAlignment(0) = 6
     .TEXT = "Nro"
     .COL = 1
     .ColWidth(1) = 1500
     .ColAlignment(1) = 1
     .TEXT = "Codigo"
     .COL = 2
     .ColWidth(2) = 1000
     .ColAlignment(2) = 1
     .TEXT = "Sub-Cod"
     .COL = 3
     .ColWidth(3) = 3000
     .ColAlignment(3) = 1
     .TEXT = "Descripcion"
     .COL = 4
     .ColWidth(4) = 800
     .ColAlignment(4) = 6
     .TEXT = "Nro.Int"
     .COL = 5
     .ColWidth(5) = 800
     .ColAlignment(5) = 6
     .TEXT = "Pzas/M."
     .COL = 6
     .ColWidth(6) = 900
     .ColAlignment(6) = 6
     .TEXT = "Cant"
     .COL = 7
     .ColWidth(7) = 900
     .ColAlignment(7) = 6
     .TEXT = "Enviar"
   End With
   '
   TEXTOENBRU$ = Clipboard.GetText
   Call SAVEFILE(LUCOM$ + "CARTEMBA.TXT", TEXTOENBRU$)
   '
   DELMAIL% = 0
   CARTEMBAR07.Text1.Visible = False
   If InStr(TEXTOENBRU$, "Entrega Número:") > 0 Then
      DELMAIL% = 1
      CARTEMBAR07.Text1.Visible = True
      Call BAJADELMAIL(TEXTOENBRU$)
      GoTo 50
   End If
   '
   PPXX% = InStr(TEXTOENBRU$, "Fecha:")
   FECHA$ = TRIM$(Mid$(TEXTOENBRU$, PPXX% + 7, 12))
   CARTEMBAR07.Label3 = FECHA$
   '
   PPXX% = InStr(TEXTOENBRU$, "O/S:")
   NOSE$ = TRIM$(Mid$(TEXTOENBRU$, PPXX% + 5, 15))
   CARTEMBAR07.Label1 = NOSE$
   CARTEMBAR07.Text1 = NOSE$
   '
   PPXX% = InStr(TEXTOENBRU$, "Entrega:")
   NENTRE$ = TRIM$(Mid$(TEXTOENBRU$, PPXX% + 8, 11))
   If XVALO(Left$(NENTRE$, 2)) < 10 Then NENTRE$ = "0" + NENTRE$
   CARTEMBAR07.Label2 = NENTRE$
   '
   N1% = FILEOPEN%(LUCOM$ + "CARTEMBA.TXT", 1, 128)
   COMILEC% = 0
   Do While Not EOF(N1%)
     Line Input #N1%, linea$
     If Left$(TRIM$(linea$), 16) = "Piezas por Medio" Then COMILEC% = 1
     If COMILEC% < 1 Then GoTo 19
     '
     VALEIDO$ = TRIM$(linea$)
     If Len(VALEIDO$) = 17 Then   ' ES UN CODIGO
        If IsNumeric(Left$(VALEIDO$, 5)) Then
          Codigo$ = Left$(VALEIDO$, 11)
          CIXI% = XVALO(VALOBADA("MASTER", "Codint", "Codigo LIKE '" + Codigo$ + "%' AND Status > 0 "))
          '
          If CIXI% < 1 Then
             Call MENSERR(24, "Atención: Código '" + Codigo$ + "' no Registrado en Base de Datos FLEXOFT.\Resultados Imprevisibles - Consulte.")
          End If
          '
          'CIXI% = CODINT%(CODIGO$)
          SUCOD$ = TRIM$(Mid$(VALEIDO$, 13))
          '
          Line Input #N1%, linea$
          NUINT& = XVALO(TRIM$(linea$))
          '
          Line Input #N1%, linea$
          PIPORME& = XVALO(TRIM$(linea$))
          '
          Line Input #N1%, linea$
          CANPIE& = XVALO(TRIM$(linea$))
          '
          Line Input #N1%, linea$
          KOK$ = SAFETEXT$(TRIM$(linea$))
          '
          Line Input #N1%, linea$
          PIENVI& = XVALO(TRIM$(linea$))
          '
          JJ& = JJ& + 1
          With CARTEMBAR07.GRID
             If .Rows <= JJ& Then .Rows = JJ& + 1
             .Row = JJ&
             .COL = 0
             .TEXT = JJ&
             .COL = 1
             .TEXT = Codigo$
             .COL = 2
             .TEXT = SUCOD$
             .COL = 3
             .TEXT = DESCRIT0$(CIXI%)
             .COL = 4
             .TEXT = NUINT&
             .COL = 5
             .TEXT = PIPORME&
             .COL = 6
             .TEXT = CANPIE&
             .COL = 7
             .TEXT = PIENVI&
          End With
          
          'Stop
        End If
     End If
     '
19 Loop
   Close #N1%
   N1% = 0
   '
50 CARTEMBAR07.Show 1
   '
End Sub
'
Sub BAJADELMAIL(TEXTOENBRU$)
   '
   PPXX% = InStr(TEXTOENBRU$, "Fecha de Entrega:")
   FECHA$ = TRIM$(Mid$(TEXTOENBRU$, PPXX% + 18, 12))
   If Mid$(FECHA$, 5, 1) = "-" Then
      FECHA$ = Mid$(FECHA$, 9, 2) + "-" + Mid$(FECHA$, 6, 2) + "-" + Left$(FECHA$, 4)
   End If
   CARTEMBAR07.Label3 = FECHA$
   '
   HOII% = HOY(HOS$)
   HHH$ = Mid$(HOS$, 7, 2) + Mid$(HOS$, 4, 2) + Mid$(HOS$, 1, 2)
   NOSE$ = "MAI-" + HHH$
   'CARTEMBAR07.Label1 = NOSE$
   '
   PPXX% = InStr(TEXTOENBRU$, "Entrega Número:")
   NENTRE$ = TRIM$(Mid$(TEXTOENBRU$, PPXX% + 15, 4))
   If XVALO(Left$(NENTRE$, 2)) < 10 Then NENTRE$ = "0" + NENTRE$
   NENTRE$ = NENTRE$ + " de 12"
   CARTEMBAR07.Label2 = NENTRE$
   '
   N1% = FILEOPEN%(LUCOM$ + "CARTEMBA.TXT", 1, 128)
   COMILEC% = 0
   Do While Not EOF(N1%)
     Line Input #N1%, linea$
     If Left$(TRIM$(linea$), 15) = "Número de Parte" Then COMILEC% = 1
     If COMILEC% < 1 Then GoTo 19
     '
     VALEIDO$ = TRIM$(linea$)
     If IsNumeric(Left$(VALEIDO$, 5)) Then
        '
        Codigo$ = Left$(VALEIDO$, 5) + "-" + Mid$(VALEIDO$, 6, 5)
        CIXI% = XVALO(VALOBADA("MASTER", "Codint", "Codigo LIKE '" + Codigo$ + "%' AND Status > 0 "))
        '
        If CIXI% <> 0 Then
           SUCOD$ = TRIM$(Mid$(VALEIDO$, 11, 4))
           '
           RESTO$ = TRIM$(Mid$(VALEIDO$, 16))
           PPXX% = InStr(RESTO$, " "): If PPXX% < 1 Then PPXX% = 1 + Len(RESTO$)
           NUINT& = XVALO(TRIM$(Left$(RESTO$, PPXX% - 1)))
           RESTO$ = TRIM$(Mid$(RESTO$, PPXX% + 1))
           '
           PPXX% = InStr(RESTO$, " "): If PPXX% < 1 Then PPXX% = 1 + Len(RESTO$)
           PIPORME& = XVALO(TRIM$(Left$(RESTO$, PPXX% - 1)))
           RESTO$ = TRIM$(Mid$(RESTO$, PPXX% + 1))
           '
           PPXX% = InStr(RESTO$, " "): If PPXX% < 1 Then PPXX% = 1 + Len(RESTO$)
           CANCAJ& = XVALO(TRIM$(Left$(RESTO$, PPXX% - 1)))
           '
           CANPIE& = CANCAJ& * PIPORME&
           KOK$ = ""
           PIENVI& = CANPIE&
           '
           If PIENVI& > 0 Then
              JJ& = JJ& + 1
              With CARTEMBAR07.GRID
                 If .Rows <= JJ& Then .Rows = JJ& + 1
                 .Row = JJ&
                 .COL = 0
                 .TEXT = JJ&
                 .COL = 1
                 .TEXT = Codigo$
                 .COL = 2
                 .TEXT = SUCOD$
                 .COL = 3
                 .TEXT = DESCRIT0$(CIXI%)
                 .COL = 4
                 .TEXT = NUINT&
                 .COL = 5
                 .TEXT = PIPORME&
                 .COL = 6
                 .TEXT = CANPIE&
                 .COL = 7
                 .TEXT = PIENVI&
              End With
           End If
        End If
     End If
     '
19 Loop
   Close #N1%
   N1% = 0
   '
End Sub
'
Private Sub Command4_Click()
   AREXE$ = VALCONTROL$(0, "APORTOYO", "DESTAPLI")
   LINCOMMA$ = VALCONTROL$(0, "APORTOYO", "LINCOMMA")
   On Error Resume Next
   NX1 = Shell(AREXE$ + " " + LINCOMMA$, vbNormalFocus)
   If Err Then
      MsgBox "Imposible Ejecutar '" + AREXE$ + " " + LINCOMMA$ + "'"
   End If
   On Error GoTo 0
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
   Label1 = ""
   Call COPYSTRU("INDIPRES", "INCARTEM")
   '
   FEXI% = FECHANUM("15/05/09") ' PARA DESCARTAR PENDIENTES ANTERIORES AL 15 DE MAYO
   FLIMITE = FETEXCOR1$(FEXI%)  ' PARA QUE TOME DESDE 15/05/09 EN ADELANTE
   '
   CONDI$ = "DOPRICOR = ''"
   CONDI$ = CONDI$ + " AND (PIE_ENVIAR > CANT_ENTRE  OR CANT_ENTRE IS NULL)"
   CONDI$ = CONDI$ + " AND FECHENTRE >= '" & FLIMITE & "'" ' 15/05/09 EN ADELANTE
   Call CARGALISEL("INCARTEM", "DETACARTEM", "FechEntre/D14; Num_Os/A16; Num_Entre/A16", CONDI$ & " ORDER BY FECHENTRE, NUM_OS, NUM_ENTRE", "DISTINCT")
   '
10 Call SELECHO("INCARTEM/Indice de Cartas de Embarque")
   If VALACT1$("INCARTEM") = "" Then GoTo 99
   '
   FESEL% = FECHANUM(VALACT1$("INCARTEM"))
   NOSE$ = TRIM$(Left$(VALACT2$("INCARTEM"), 16))
   NUENTRE$ = TRIM$(Mid$(VALACT2$("INCARTEM"), 17, 16))
   CEMBASE$ = AJUSTI$(VALACT1$("INCARTEM"), 16) + VALACT2$("INCARTEM")
   '
   condi0$ = "Num_Os = '" & NOSE$ & "' AND Num_Entre = '" & NUENTRE$ & "' AND FechEntre = '" & FETEXCOR1$(FESEL%) & "'"
   YAESTA% = CantRegistros("ENCACARTEM", condi0$)
   If YAESTA% < 1 Then
      Call MENSERR(24, "Carta de Embarque no Registrada.")
      GoTo 10
   End If
   '
   CONDI$ = condi0$ + " AND (Dopricor = '' OR Dopricor IS NULL OR Pie_Enviar > Cant_Entre ) "
   CANREPRO% = CantRegistros("DETACARTEM", CONDI$)
   If CANREPRO% < 1 Then
      TTXX$ = "Carta de Embarque ya Totalmente Despachada.\  \Imposible Volver a Despachar."
      GoTo 99
   End If
   '
   CODFOR$ = TRIM$(UCase$(CONTROL$("", "FOREMIFA")))
   TOTITFO& = Val(ATRIFORM$(CODFOR$, "MAX-ITEM"))
   If CANREPRO% > TOTITFO& Then
      Call COMUNI("Los Items Pendientes de esta Carta de Embarque\Superan el Máximo Permitido para el Remito.\   \Los Excedentes quedan Pendientes\para el Próximo Remito.")
   End If
   '
   CONDI$ = CONDI$ + " ORDER BY NorItCa ASC "
   JJ& = 0
   REBLA$ = REGBLAN$("BOLREDET")
   Set DECARTE = FLEXCONN.Execute("SELECT * FROM DETACARTEM WHERE " + CONDI$)
   With DECARTE
      If Not (.BOF And .EOF) Then
         Do While Not .EOF
            XX$ = REBLA$
            CIXI% = XVALO(!cod_inte)
            CANTU# = XVALO(!Pie_Enviar) - XVALO(!Cant_Entre)
            If CANTU# < 0.05 Then GoTo 20
            UM$ = Unimed$(CIXI%)
            SUCOD$ = SAFETEXT$(!SubCod)
            NORIX% = XVALO(!NorItCa)
            Call REPLA(XX$, MKI$(CIXI%), 83, 2)
            Call REPLA(XX$, UM$, 85, 4)
            Call REPLA(XX$, MKD$(CANTU#), 89, 8)
            Call REPLA(XX$, SUCOD$, 97, 8)
            Call REPLA(XX$, MKI$(NORIX%), 187, 2)
            JJ& = JJ& + 1
            Call GRAREG("!BOLREDET", XX$, JJ&)
            If JJ& >= TOTITFO& Then GoTo 29
20        .MoveNext
         Loop
      End If
   End With
   '
29 DECARTE.Close
   Set DECARTE = Nothing
   Call SETULTREG("!BOLREDET", JJ&)
   '
   Label1 = CEMBASE$
   '
   FOREMIT07.Text22 = NOSE$
   FOREMIT07.Text16(0) = "ENTREGA Nº " & NUENTRE$
   FOREMIT07.Text22.Enabled = False
   FOREMIT07.Text16(0).Enabled = False
   '
99 Command5.Enabled = True
   If Label1 <> "" Then
      Hide
   End If
   '
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   
   FORIPRE$ = TRIM$(CONTROL$("", "FORMVALE"))
   If FORIPRE$ <> "" Then
      If ECOARCH$("INDIFRM", FORIPRE$) <> "" Then
         Call COPYSTRU("INDIPRES", "INCARTEM")
         '
         FEXI% = FECHANUM("15/05/09") ' PARA DESCARTAR PENDIENTES ANTERIORES AL 15 DE MAYO
         FLIMITE = FETEXCOR1$(FEXI%)  ' PARA QUE TOME DESDE 15/05/09 EN ADELANTE
         '
         CONDI$ = "DOPRICOR = ''"
         CONDI$ = CONDI$ + " AND (PIE_ENVIAR > CANT_ENTRE  OR CANT_ENTRE IS NULL)"
         CONDI$ = CONDI$ + " AND FECHENTRE >= '" & FLIMITE & "'" ' 15/05/09 EN ADELANTE
         Call CARGALISEL("INCARTEM", "DETACARTEM", "FechEntre/D14; Num_Os/A16; Num_Entre/A16", CONDI$ & " ORDER BY FECHENTRE, NUM_OS, NUM_ENTRE", "DISTINCT")
         '
10       Call SELECHO("INCARTEM/Indice de Cartas de Embarque")
         If VALACT1$("INCARTEM") = "" Then GoTo 99
         '
         FESEL% = FECHANUM(VALACT1$("INCARTEM"))
         NOSE$ = TRIM$(Left$(VALACT2$("INCARTEM"), 16))
         NUENTRE$ = TRIM$(Mid$(VALACT2$("INCARTEM"), 17, 16))
         CEMBASE$ = AJUSTI$(VALACT1$("INCARTEM"), 16) + VALACT2$("INCARTEM")
         '
         condi0$ = "Num_Os = '" & NOSE$ & "' AND Num_Entre = '" & NUENTRE$ & "' AND FechEntre = '" & FETEXCOR1$(FESEL%) & "'"
         YAESTA% = CantRegistros("ENCACARTEM", condi0$)
         If YAESTA% < 1 Then
            Call MENSERR(24, "Carta de Embarque no Registrada.")
            GoTo 10
         End If
         '
         CONDI$ = condi0$ + " ORDER BY NorItCa ASC "
         '
         JJ& = 0
         Set DECARTE = FLEXCONN.Execute("SELECT * FROM DETACARTEM WHERE " + CONDI$)
         '
         FEX = FESEL%: FECHDOC$ = FECHATEX$(FEX)
         DOCUORIG$ = TRIM$(NUENTRE$) + " - " + NOSE$
         TIPODOC$ = "Vale de Preparacion de Entregas"
         '
         CODTABU1$(1) = "COD-MAT"
         CODTABU1$(2) = "REF-MAT1"
         CODTABU1$(3) = "CANTIDAD"
         CODTABU1$(4) = "MEDIDA-1"
         CODTABU1$(5) = "ORD-ITEM"
         CACOLTAB1% = 5
         CAITAB1% = 0
         '
         With DECARTE
            If Not (.BOF And .EOF) Then
               Do While Not .EOF
                  JJ& = JJ& + 1
                  CAITAB1% = JJ&
                    CIXI% = XVALO(!cod_inte)
                    SUCOD$ = SAFETEXT$(!SubCod)
                    CODE$ = TRIM$(Left$(CODEXT$(CIXI%), 11)) + "-" + SUCOD$
                  TETABU1$(1, JJ&) = CODE$
                    NINTER& = XVALO(!Num_Inter)
                  TETABU1$(2, JJ&) = TRIM$(Str$(NINTER&))
                    PIPORCA& = XVALO(!Pie_Medio)
                  TETABU1$(3, JJ&) = TRIM$(Str$(PIPORCA&))
                    CAJPEDI& = 1
                    CANPIEZAS& = XVALO(!Pie_Enviar)
                    If PIPORCA& > 0 Then CAJPEDI& = CANPIEZAS& / PIPORCA&
                  TETABU1$(4, JJ&) = TRIM$(Str$(CAJPEDI&))
                  TETABU1$(5, JJ&) = TRIM$(Str$(JJ&))
               .MoveNext
               Loop
            End If
         End With
         '
29       DECARTE.Close
         Set DECARTE = Nothing
         Call PRINTDOC("FORMVALE")
         '
         GoTo 10
      End If
   End If
   '
   Call MENSERR(24, "Falta Formulario de Impresión")
   '
99 Command6.Enabled = True
End Sub

Private Sub Command8_Click()
  Call FINAL("?CARTEMBAR")
End Sub

Private Sub Command14_Click()
  Call FINAL("?CAEMBAREM")
End Sub


Sub MUESTREMI(NUREMIT$)

    NUDOCU$ = NUREMIT$
    TIDOCUM$ = "Remito Nro."
    CONDI$ = "DOPRICOR = '" & NUREMIT$ & "'"
    NUDOCU$ = VALOBADA("DETACARTEM", "DOPRILAR", CONDI$, "DISTINC")
    If NUDOCU$ = "" Then Exit Sub
    Screen.MousePointer = 11
    '
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
         XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
         If InStr(XX$, TIDOCUM$) > 4 Then
           PPXX% = InStr(XX$, "Nro.")
           If PPXX% > 0 Then
             XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
           End If
           If InStr(XX$, NUDOCU$) > 4 Then
             DOCUNU& = CVS(Left$(XX$, 4))
             GoTo 5
           End If
         End If
    Next UI&
    '
    Screen.MousePointer = 1
    Call COMUNI("No se Encontró el Documento Impreso\Correspondiente al Remito " & NUDOCU$ & ".")
   '
5   Call MUESTRADOC(DOCUNU&) 'PERMITE VER LOS DOCUMENTOS ANTIGUOS POR PANTALLA
   
End Sub

Private Sub Form_Load()
    '
    Call APLICACIONSKINS(Me, Picture8)
    '
End Sub
Sub TEST()
'AL AGREGAR UN CODIGO A LA  GRILLA
'EST_PROD SOLO EN ESTA
'NO SE GRABA SOLO EN SELECT
'EST_PROD_PRECO SOLO EN SELECT
'EST_PROD_CPL SOLO EN SELECT

'Select  P.CD_PROD,  P.DS_PROD, (CASE WHEN ISNULL(P.VLR_CUSTO_GRP, 0) > 0 THEN P.VLR_CUSTO_GRP      ELSE CPL.VLR_CUSTO END) AS VLR_CUSTO_GRP,  CASE WHEN V.VLR_OFERTA  > 0 THEN V.VLR_OFERTA ELSE 0  END AS VLR, VLR_TABELA, EST_PROD_FABRIC.NM_FABRIC
'From  EST_PROD P
'INNER JOIN EST_PROD_FABRIC   ON P.CD_FABRIC = EST_PROD_FABRIC.CD_FABRIC
'LEFT JOIN EST_PROD_PRECO V ON   P.CD_EMP = V.CD_EMP AND   P.CD_PROD =
'V.CD_PROD AND   V.CD_FILIAL = 1
'
' LEFT JOIN EST_PROD_CPL CPL ON   V.CD_EMP = CPL.CD_EMP AND   V.CD_PROD = CPL.CD_PROD AND   V.CD_FILIAL = CPL.CD_FILIAL
'
'Where p.CD_EMP = 1 And p.CD_PROD = 50572
'
'AL grabar
'
'Update EST_PROD_TABLOIDE
' SET DT_INICIO = '2020-02-01',DT_FIM = '2022-12-11',DS_DESCRICAO = 'VEREJINHO/HOSPITALAR',CD_USU = 1,DT_ULT_ALT = '2022-02-17',FLAG_VIGENCIA = 0
' Where CD_EMP = 1 And CD_TABLOIDE = 162
'
'BORRA DE TABLA EST_PROD_TABLOIDE_FILIAL
'DELETE FROM EST_PROD_TABLOIDE_FILIAL WHERE  CD_EMP = 1 AND CD_TABLOIDE = 162
'
'UPDATE EN EST_PROD_TABLOIDE_CPL
'UPDATE EST_PROD_TABLOIDE_CPL SET STS_PROD = 1
'Where CD_TABLOIDE = 162
'
'SIGUE:
'
'ANTES DE HACER LOS UPDATE VALIDA QUE EXISTA LA COMIBINACION
'CD_EMP = 1 And CD_TABLOIDE = 162 And CD_PROD = 34791
'
'Update EST_PROD_TABLOIDE_CPL
' SET CD_EMP = 1,CD_TABLOIDE = 162,CD_PROD = 34791,VLR_GRP_PROD =  60,STS_PROD = 0
'Where CD_EMP = 1 And CD_TABLOIDE = 162 And CD_PROD = 34791
'
'REPITE PARA CADA ITEM
'UPDATE EST_PROD_TABLOIDE_CPL SET CD_EMP = 1,CD_TABLOIDE = 162,CD_PROD = 158751,VLR_GRP_PROD =  69.99,STS_PROD = 0
' Where CD_EMP = 1 And CD_TABLOIDE = 162 And CD_PROD = 158751
'
'SI NO EXISTE
'
'INSERT INTO EST_PROD_TABLOIDE_CPL(CD_EMP,CD_TABLOIDE,CD_PROD,VLR_GRP_PROD,STS_PROD) VALUES (1,162,50572, 70,0)
'
'
'AL FINAL
'
'LO HACE PARA TODAS LAS FILIALES DE A UNA
'INSERT INTO EST_PROD_TABLOIDE_FILIAL(CD_EMP,CD_TABLOIDE,CD_FILIAL) VALUES (1,162,1)
'
'
'
'
'
'


End Sub

