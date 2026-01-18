VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form GEFLXPRFS1 
   BackColor       =   &H0095C5E1&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Genera Archivos de Extracción de Controladores"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   6810
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox MARCOBARRA 
      Appearance      =   0  'Flat
      BackColor       =   &H002B9973&
      ForeColor       =   &H80000008&
      Height          =   200
      Left            =   5490
      ScaleHeight     =   165
      ScaleWidth      =   1095
      TabIndex        =   12
      Top             =   1560
      Width           =   1125
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H0048CC9D&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1770
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   2535
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   1455
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   2566
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ListBox List3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      ItemData        =   "GEFLXPRFS1.frx":0000
      Left            =   5480
      List            =   "GEFLXPRFS1.frx":0002
      TabIndex        =   5
      Top             =   1920
      Width           =   1170
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      Left            =   120
      TabIndex        =   3
      Top             =   3480
      Width           =   6525
   End
   Begin VB.CommandButton Command1 
      Height          =   435
      Left            =   5520
      Picture         =   "GEFLXPRFS1.frx":0004
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Inicia Exploraciónde la Red"
      Top             =   120
      Width           =   1170
   End
   Begin VB.ListBox List1 
      Height          =   1280
      IntegralHeight  =   0   'False
      Left            =   240
      TabIndex        =   1
      Top             =   3840
      Visible         =   0   'False
      Width           =   5160
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1710
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   2430
   End
   Begin VB.Label Label5 
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
      Left            =   315
      TabIndex        =   9
      Top             =   2140
      Width           =   5055
   End
   Begin VB.Label Label4 
      Caption         =   "Controladores Actualizados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3915
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   7
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   5480
      TabIndex        =   6
      Top             =   2985
      Width           =   1170
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   5480
      TabIndex        =   4
      Top             =   2640
      Width           =   1170
   End
End
Attribute VB_Name = "GEFLXPRFS1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   '
   Dim NOMARCH$(), FEARCH#(), RUTARCH$()
   '
   MSF.ColWidth(0) = 4000
   ULODAX$ = "\\SRVSQL01\discof\DOCU_TEC_SYSFLOM\SYS_BAS\FLEXPRFS"
   ULODAY$ = "F:\DOCU_TEC_SYSFLOM\SYS_BAS\FLEXPRFS"
   For KKI% = 48 To 57
      On Error Resume Next
      Kill ULODAY$ + "\RFSX\TRAN" + Chr$(KKI%) + "*.RFS"
      On Error GoTo 0
   Next KKI%
   '
   For LETRA% = 67 To 90
      RUTAX$ = Chr$(LETRA%) + ":\"
      If LETRA% > 67 Then MSF.Rows = MSF.Rows + 1
      MSF.TextMatrix(MSF.Rows - 1, 0) = RUTAX$
   Next LETRA%
   '
   Label5 = "Explorando las Carpetas del Disco Servidor (\\ServerNew) ..."
   Do While ILIST1& < MSF.Rows
      RUTAX$ = MSF.TextMatrix(ILIST1&, 0)
      On Error Resume Next
      NOMBREX$ = Dir(RUTAX$, 16) ' Recupera la primera entrada.
      If Err Then
         On Error GoTo 0
         GoTo 6
      End If
      On Error GoTo 0
      Do While NOMBREX$ <> "" ' Inicia el bucle.
        ' Ignora el directorio actual y el que lo abarca.
        If NOMBREX$ <> "." And NOMBREX$ <> ".." Then
           ' Utiliza comparación a nivel de bits para asegurarse de que MiNombre es un directorio.
           If (GetAttr(RUTAX$ & NOMBREX$) And vbDirectory) = vbDirectory Then
              If UCase$(RUTAX$ + NOMBREX$) <> ULODAY$ Then ' EXCLUYE DIRECTORIO TRABAJO
                 MSF.Rows = MSF.Rows + 1
                 MSF.TextMatrix(MSF.Rows - 1, 0) = RUTAX$ + NOMBREX$ + "\"
                 'List1.ListIndex = List1.ListCount - 1
              End If
           End If ' solamente si representa un directorio.
        End If
        NOMBREX$ = Dir ' Obtiene siguiente entrada.
        OPENFORMS = DoEvents
      Loop
6     ILIST1& = ILIST1& + 1
      Label1 = ILIST1&: Label2 = MSF.Rows
   Loop
   '
   '
   '
   List2.Clear
   TICO$ = "RFS": VUELTA% = 1
' CARGA EN MEMORIA LOS RFS / PRF ACTUALES
10 Screen.MousePointer = 11
   Label5 = "Carga en Memoria los " + TICO$ + " Actuales ..."
   DIRFSX$ = ULODAX$ + "\" + TICO$ + "X"
   Dir1.Path = DIRFSX$
   File1.Path = Dir1.Path
   File1.Pattern = "*." + TICO$
   CARCHI% = 0: MAXARCH% = 0
   OPENFORMS = DoEvents
   '
   FIN& = File1.ListCount
   For U& = 1 To File1.ListCount
      Call TRACE(20, U&, FIN&)
      FINAX$ = UCase$(File1.List(U& - 1))
      If VUELTA% = 1 Then
          If Left$(FINAX$, 4) = "TRAN" Then
              If Asc(Mid$(FINAX$, 5, 1)) < 58 Then
                  GoTo 17
              End If
          End If
      End If
      '
      If Asc(Left$(FINAX$, 1)) >= 48 Then
        If Asc(Left$(FINAX$, 1)) <= 90 Then
          LOARX = 0: FIDAT# = 0
          On Error Resume Next
          LOARX = FileLen(DIRFSX$ + "\" + FINAX$)
          FIDAT# = FileDateTime(DIRFSX$ + "\" + FINAX$)
          If LOARX > 0 Then
            If FIDAT# > 0 Then
              CARCHI% = CARCHI% + 1
              If CARCHI% > MAXARCH% Then
                 MAXARCH% = MAXARCH% + 1024
                 ReDim Preserve NOMARCH$(MAXARCH%), FEARCH#(MAXARCH%), RUTARCH$(MAXARCH%)
              End If
              FEARCH#(CARCHI%) = FIDAT#
              NOMARCH$(CARCHI%) = FINAX$
              RUTARCH$(CARCHI%) = ""
            End If
          End If
          On Error GoTo 0
        End If
      End If
17 Next U&
   '
20 Label5 = "Revisa los " + TICO$ + " Nuevos contra los Actuales ..."
   For ILISTA& = 1 To MSF.Rows
     Call TRACE(20, ILISTA&, FIN&)
     On Error Resume Next
     File1.Path = MSF.TextMatrix(ILISTA& - 1, 0)
     If Err Then
        On Error GoTo 0
        GoTo 29
     End If
     On Error GoTo 0
     If UCase$(File1.Path) = UCase$(DIRFSX$) Then GoTo 29
     DoEvents
     File1.Pattern = "*." + TICO$ + ""
     For KKI% = 1 To File1.ListCount
        FINAME$ = TRIM$(UCase$(File1.List(KKI% - 1)))
        If FINAME$ <> "" Then
           If VUELTA% = 1 Then
               If Left$(FINAME$, 4) = "TRAN" Then
                   If Asc(Mid$(FINAME$, 5, 1)) < 58 Then
                       GoTo 28
                   End If
               End If
           End If
           PASOX$ = MSF.TextMatrix(ILISTA& - 1, 0)
           If Right$(PASOX$, 1) <> "\" Then PASOX$ = PASOX$ + "\"
           ACOMPLE$ = PASOX$ + FINAME$
           If FileLen(ACOMPLE$) > 0 Then
              LARCHI& = FileLen(ACOMPLE$)
              Label1.Caption = Str$(Val(Label1.Caption) + LARCHI&)
              DoEvents
              FECHAC# = FileDateTime(ACOMPLE$)
              If FECHAC# <= CDbl(Now) Then
                 For KKJ% = 1 To CARCHI%
                    If FINAME$ = NOMARCH(KKJ%) Then
                       If FECHAC# > FEARCH#(KKJ%) Or (FECHAC# = FEARCH#(KKJ%) And InStr(ACOMPLE$, "\" + TICO$ + "X") > 0) Then
                          FEARCH#(KKJ%) = FECHAC#
                          RUTARCH$(KKJ%) = MSF.TextMatrix(ILISTA& - 1, 0)
                       End If
                       GoTo 28
                    End If
                 Next KKJ%
                 '
                 If Asc(Left$(FINAME$, 1)) >= 48 Then
                    If Asc(Left$(FINAME$, 1)) <= 90 Then
                       CARCHI% = CARCHI% + 1
                       If CARCHI% > MAXARCH% Then
                          MAXARCH% = MAXARCH% + 1024
                          ReDim Preserve NOMARCH$(MAXARCH%), FEARCH#(MAXARCH%), RUTARCH$(MAXARCH%)
                       End If
                       If FECHAC# > CDbl(Now) Then FECHAC# = 0
                       NOMARCH$(CARCHI%) = FINAME$
                       FEARCH#(CARCHI%) = FECHAC#
                       RUTARCH$(CARCHI%) = MSF.TextMatrix(ILISTA& - 1, 0)
                    End If
                 End If
              End If
           End If
        End If
28   Next KKI%
29 Next ILISTA&
   '
30 FIN& = CARCHI%
   Label5 = "Arma la lista de Actualización de los " + TICO$ + " ..."
   For KKI% = 1 To CARCHI%
      KKL& = KKI%: Call TRACE(20, KKL&, FIN&)
      If TRIM$(RUTARCH$(KKI%)) <> "" Then
         FERFSX# = 0
         On Error Resume Next
         FERFSX# = CDbl(FileDateTime(ULODAX$ + "\" + TICO$ + "X\" + NOMARCH(KKI%)))
         On Error GoTo 0
         If FEARCH#(KKI%) > FERFSX# Then
            DATU$ = Format$(FEARCH#(KKI%), "DD-MM-YY")
            DATR$ = Format$(FERFSX#, "DD-MM-YY")
            On Error Resume Next
            FINUE$ = LOADFILE$(RUTARCH$(KKI%) + NOMARCH$(KKI%))
            FIVIE$ = LOADFILE$(ULODAX$ + "\" + TICO$ + "X\" + NOMARCH$(KKI%))
            If FINUE$ <> FIVIE$ Then
               If FINUE$ <> "" Then
                  AAA = DATU$ + "   " + RUTARCH$(KKI%) + NOMARCH$(KKI%) + "   " + DATR$
                  List2.AddItem AAA
                  FileCopy RUTARCH$(KKI%) + NOMARCH$(KKI%), ULODAX$ + "\" + TICO$ + "X\" + NOMARCH$(KKI%)
               End If
            End If
            On Error GoTo 0
         End If
      End If
   Next KKI%
   '
   '
   '
50 If VUELTA% = 1 Then
      For KKI% = 1 To CARCHI%
         NOMARCH$(KKI%) = ""
         FEARCH#(KKI%) = 0
         RUTARCH$(KKI%) = ""
      Next KKI%
      CARCHI% = 0
      VUELTA% = 2
      TICO$ = "PRF"
      GoTo 10
   End If
   Screen.MousePointer = 1
   '
   If List2.ListCount < 1 Then
      MsgBox "FIN SIN ACTUALIZACIONES"
      Exit Sub
   End If
   CANACTU% = List2.ListCount
   '
98 Screen.MousePointer = 11
   '
   Label3 = 12
   'Call DOSCOMMAND("PKZIP " + ULODAX$ + "\FLEXRFS " + ULODAX$ + "\RFSX\*.RFS")
   'Call DOSCOMMAND("PKZIP " + ULODAX$ + "\FLEXPRF " + ULODAX$ + "\PRFX\*.PRF")
   '
99 On Error Resume Next
   Kill ULODAX$ + "\FLEXPRFS.FLX"
   On Error GoTo 0
   '
   Dim BUFLE As String * 128
   Dim BUFGRA As String * 160
   NX% = FreeFile
   Open ULODAX$ + "\FLEXPRFS.FLX" For Random Access Read Write Shared As #NX% Len = 160
   JJ& = 0
   '
   Label3 = 13
   MSF.Clear
   File1.Path = ULODAX$ + "\RFSX"
   DoEvents
   File1.Pattern = "*.RFS"
   '
   Label3 = 14
   FIN& = File1.ListCount
   For KKI% = 1 To File1.ListCount
     KKL& = KKI%
     Call TRACE(20, KKL&, FIN&)
     FINAME$ = TRIM$(UCase$(File1.List(KKI% - 1)))
     If FINAME$ <> "" Then
       For JKI% = 1 To List3.ListCount
         If List3.List(JKI% - 1) > FINAME$ Then
           List3.AddItem FINAME$, JKI% - 1
           GoTo 14
         End If
       Next JKI%
       List3.AddItem FINAME$
     End If
14 Label2 = List3.ListCount
   Next KKI%
   '
   Label3 = 15
   File1.Path = ULODAX$ + "\PRFX"
   DoEvents
   File1.Pattern = "*.PRF"
   '
   Label3 = 16
   FIN& = File1.ListCount
   For KKI% = 1 To File1.ListCount
     KKL& = KKI%
     Call TRACE(20, KKL&, FIN&)
     FINAME$ = TRIM$(UCase$(File1.List(KKI% - 1)))
     If FINAME$ <> "" Then
       For JKI% = 1 To List3.ListCount
         If List3.List(JKI% - 1) > FINAME$ Then
           List3.AddItem FINAME$, JKI% - 1
           GoTo 16
         End If
       Next JKI%
       List3.AddItem FINAME$
     End If
16 Label2 = List3.ListCount
   Label2.Refresh
   Next KKI%
   '
   Label3 = 17
   Screen.MousePointer = 11
   FIN& = List3.ListCount
   For KKI% = 1 To List3.ListCount
     KKL& = KKI%
     Call TRACE(20, KKL&, FIN&)
     FINAME$ = TRIM$(UCase$(List3.List(KKI% - 1)))
     If FINAME$ <> "" Then
       If Left$(FINAME$, 1) <> "." Then
         CARPETRA$ = ULODAX$ + "\RFSX"
         If InStr(FINAME$, ".PRF") > 0 Then
           CARPETRA$ = ULODAX$ + "\PRFX"
         End If
         '
         ACOMPLE$ = CARPETRA$ + "\" + FINAME$
         LARCHI& = 0
         On Error Resume Next
         LARCHI& = FileLen(ACOMPLE$)
         On Error GoTo 0
         CARELE& = Int((LARCHI + 127) / 128)
         If LARCHI& > 0 Then
           N1% = FreeFile
           Open ACOMPLE$ For Random Access Read As #N1% Len = 128
           For II& = 1 To CARELE&
             Get #N1%, II&, BUFLE$
             BUFGRA$ = Space$(160)
             Mid$(BUFGRA$, 1, 32) = FINAME$
             Mid$(BUFGRA$, 33, 128) = BUFLE$
             JJ& = JJ& + 1
             Put #NX%, JJ&, BUFGRA$
           Next II&
           Close #N1%
         End If
       End If
     End If
   Next KKI%
   '
   Screen.MousePointer = 1
   List3.Visible = True
   Close #NX%
   '
   MsgBox "Proceso Completo con " & CANACTU% & " Actualizaciones. "
   '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Close: End

End Sub

