VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form ADJUNARCH 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adjuntar Archivos"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   9675
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
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
      Height          =   525
      Left            =   8865
      Picture         =   "ADJUNARCH.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Prepara para Subir a Sitio FLEXOFT"
      Top             =   1270
      Width           =   700
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   8865
      Picture         =   "ADJUNARCH.frx":0E42
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   200
      Width           =   700
   End
   Begin VB.CommandButton Command2 
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
      Height          =   535
      Left            =   8865
      Picture         =   "ADJUNARCH.frx":0F8C
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Adjuntar / Procesar"
      Top             =   740
      Width           =   700
   End
   Begin VB.TextBox Text2 
      Height          =   435
      Left            =   420
      TabIndex        =   13
      Text            =   "Text2"
      Top             =   6405
      Width           =   3165
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Lista de Archivos Adjuntos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   1695
      Left            =   105
      TabIndex        =   8
      Top             =   105
      Width           =   8625
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Left            =   210
         TabIndex        =   9
         Top             =   350
         Width           =   8205
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Selección Archivo Adjunto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3690
      Left            =   105
      TabIndex        =   0
      Top             =   1890
      Width           =   9465
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Filtros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   315
         TabIndex        =   14
         Top             =   2460
         Width           =   3165
         Begin VB.OptionButton Option4 
            BackColor       =   &H00C0FFFF&
            Caption         =   "FRM"
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
            Left            =   2325
            TabIndex        =   19
            Top             =   315
            Width           =   645
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00C0FFFF&
            Caption         =   "RFS"
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
            Left            =   1620
            TabIndex        =   18
            Top             =   315
            Width           =   645
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0FFFF&
            Caption         =   "PRF"
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
            Left            =   915
            TabIndex        =   17
            Top             =   315
            Width           =   645
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "EXE"
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
            Left            =   210
            TabIndex        =   16
            Top             =   315
            Value           =   -1  'True
            Width           =   645
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   210
            TabIndex        =   15
            Top             =   630
            Width           =   2745
         End
      End
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2460
         IntegralHeight  =   0   'False
         Left            =   3570
         TabIndex        =   12
         Top             =   420
         Width           =   5685
      End
      Begin VB.ListBox List2 
         Height          =   2205
         Left            =   5670
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   420
         Width           =   1170
      End
      Begin VB.FileListBox File1 
         Height          =   2040
         Left            =   3780
         TabIndex        =   5
         Top             =   420
         Width           =   1695
      End
      Begin VB.DirListBox Dir1 
         Height          =   1440
         Left            =   315
         TabIndex        =   2
         Top             =   945
         Width           =   3165
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   1260
         TabIndex        =   1
         Top             =   420
         Width           =   2220
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   105
         TabIndex        =   7
         Top             =   465
         Width           =   1065
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4305
         TabIndex        =   6
         Top             =   1995
         Width           =   4320
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3990
         TabIndex        =   4
         Top             =   3230
         Width           =   5265
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Carpeta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3255
         TabIndex        =   3
         Top             =   2980
         Width           =   1065
      End
   End
   Begin VB.Label Label3 
      Height          =   435
      Left            =   4200
      TabIndex        =   22
      Top             =   6405
      Width           =   3690
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   225
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   120
   End
End
Attribute VB_Name = "ADJUNARCH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   '
   Dim FINX$()
   ReDim FINX$(0)
   
   Command1.Enabled = False
   OPXX% = COMALTER%("Seleccionar Opción \\Actualizar Novedades\Actualizar Controladores")

   If OPXX% = 2 Then
       Call CREAR_INDICE
       Command1.Enabled = True
       Exit Sub
     Else
       '
       ARINDEX$ = "F:\WEBUPDATES\NOVEDAD\IUPDATES.TXT"
       If EXISTE%(ARINDEX$) > 0 Then
          FULTACT$ = LOADFILE$(ARINDEX$)
          CANREGACT& = Len(FULTACT$) / 128
          JJ& = 0
          ReDim FINX$(CANREGACT&)
          For U& = 1 To CANREGACT&
             FIRECVI$ = Mid$(FULTACT$, 128 * U& - 127, 128)
             If TRIM$(Left$(FIRECVI$, 64)) <> "" Then
               XBUS$ = TRIM$(UCase$(Left$(FIRECVI$, 64)))
               FBUS& = XVALO(Mid$(FIRECVI$, 65, 5))
               USEP$ = TRIM$(UCase$(Mid$(FIRECVI$, 75, 28)))
               For KK& = 1 To JJ&
                 If TRIM$(UCase$(Left$(FINX$(KK&), 64))) = XBUS$ Then
                   If TRIM$(UCase$(Mid$(FIRECVI$, 75, 28))) = USEP$ Then
                     If XVALO(Mid$(FINX$(KK&), 65, 5)) < FBUS& Then
                       FINX$(KK&) = FIRECVI$
                     End If
                     GoTo 8
                   End If
                 End If
               Next KK&
               JJ& = JJ& + 1
               FINX$(JJ&) = FIRECVI$
             End If
8         Next U&
       End If
       '
       FIN& = List1.ListCount
       For U& = 1 To FIN&      '
          ACOMPLE$ = UCase$(TRIM$(List1.List(U& - 1)))
          TTXX$ = ACOMPLE$
10        PPXX% = InStr(TTXX$, "\")
          If PPXX% > 0 Then
             TTXX$ = Mid$(TTXX$, PPXX% + 1)
             GoTo 10
          End If
          TXTNOMARC$ = TTXX$
          '
          TTXX$ = Format$(FileDateTime(ACOMPLE$), "dd/mm/yy  hh:mm")  'Se obtiene la fecha de la ultima modificacion
          TXTFECHA% = FECHANUM(Left$(TTXX$, 8))           'Se obtiene la fecha de la ultima modificacion
          TXTHORA% = Int(HORANUM(Mid$(TTXX$, 11, 5)))         'Se obtiene la hora de la ultima modificacion
          '
          XX$ = Space$(128)
          'Se colocan los los datos en el registro
          Call REPLA(XX$, TXTNOMARC$, 1, 64)
          Call REPLA(XX$, FORMATNUM$(TXTFECHA%, "I5"), 65, 5)
          Call REPLA(XX$, FORMATNUM$(TXTHORA%, "I4"), 71, 4)
          '
          For KK& = 1 To JJ&
             If TRIM$(Left$(FINX$(KK&), 64)) = TXTNOMARC$ Then
                FINX$(KK&) = XX$    ' esto está bien - Es para los Ejecutables - EPB 5/11/09
                GoTo 29
             End If
          Next KK&
          If JJ& >= UBound(FINX$) Then
             ReDim Preserve FINX$(JJ& + 1)
          End If
          JJ& = JJ& + 1
          FINX$(JJ&) = XX$
          '
29        ABC1 = COPYFILE%(ACOMPLE$, "F:\WEBUPDATES\NOVEDAD\" + TXTNOMARC$)
          '
       Next U&
       '
       ' ESCRIBE EN EL ARCHIVO
       FULTACT$ = ""
       For KK& = 1 To JJ&
          FULTACT$ = FULTACT$ + FINX$(KK&)
       Next KK&
       Call SAVEFILE("F:\WEBUPDATES\NOVEDAD\IUPDATES.TXT", FULTACT$)
       '
       Call MENSERR(24, "Actualización de Novedades Realizada")
       Command1.Enabled = True
   End If
   '
End Sub

Sub CREAR_INDICE()
     '
  Drive1.Drive = "F:"
  Dir1.Path = "F:\DOCU_TEC\ARMA_CD\FX_SETUP"
  Text1.TEXT = "*.*"
  '
  FIN& = List3.ListCount
  NUMARCH% = FILEOPEN%("F:\WEBUPDATES\FX_SETUP\LF_SETUP.TXT", 2, 128)

   For U& = 1 To FIN&
      LINEA$ = UCase$(TRIM$(List3.List(U& - 1)))
      NOMARCH$ = TRIM$(Left$(LINEA$, 32))
      ABC1 = COPYFILE(Dir1.Path + "\" + NOMARCH$, "F:\WEBUPDATES\FX_SETUP\" + NOMARCH$)
      Print #NUMARCH%, LINEA$
   Next U&
   '
   Close #NUMARCH%
   Call MENSERR(24, "Actualización de Controladores Realizada")

End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   '
   If List1.ListCount < 1 Then
      Call MENSERR(24, "No se han Seleccionado Archivos a Adjuntar")
      GoTo 99
   End If
   '
   HOII% = HOY(HOS$)
   AZIP0$ = "UP" + Mid$(HOS$, 7, 2) + Mid$(HOS$, 4, 2) + Mid$(HOS$, 1, 2) + "-"
   IDX% = 0
10 AZIP$ = AZIP0$ + TRIM$(Str$(IDX%))
   If EXISTE%("F:\ANOVE\" + AZIP$ + ".ZIP") > 0 Then
     IDX% = IDX% + 1
     GoTo 10
   End If
   AZIPSH$ = AZIP$
   While InStr(AZIPSH$, "\") > 0
     AZIPSH$ = Mid$(AZIPSH$, 2)
   Wend
   '
   TTXX$ = ""
   For U& = 1 To List1.ListCount
     TTXX$ = TTXX$ + Chr$(34) + ReplaSkey$(List1.List(U& - 1)) + Chr$(34) + Chr$(32)
   Next U&
  
   DESTINO$ = "F:\ANOVE\" + AZIP$ + ".ZIP"
   Label3 = "F:\ANOVE\" + AZIP$ + ".ZIP"
   On Error Resume Next
   RETURNVALUE = Shell("C:\Archivos de programa\WinZip\WZZIP.EXE " + DESTINO$ + " " & TTXX$, 4)
   If Err Then
      On Error GoTo 0
      Call MENSERR(24, "Imposible Procesar por WZZIP.EXE. Se Envía por WINZIP Normal.")
      GoTo 12
   End If
   GoTo 99
   '
   '
12 RUTAWIN$ = "C:\Archivos de programa\WinZip\WINZIP32.EXE"
   If EXISTE%(RUTAWIN$) < 1 Then
      RUTAWIN$ = "D:\Archivos de programa\WinZip\WINZIP32.EXE"
   End If
   '
   RETURNVALUE = Shell(RUTAWIN$, 4)
   If Err Then
     On Error GoTo 0
     MsgBox "No Existe " + RUTAWIN$
     GoTo 99
   End If
   On Error GoTo 0
   On Error Resume Next
   AppActivate RETURNVALUE   ' : AppActivate RETURNVALUE
   If Err Then
      On Error GoTo 0
      MsgBox "Imposible Activar WINZIP32"
      GoTo 99
   End If
   On Error GoTo 0
   SendKeys "^N", True
   SendKeys "F:\ANOVE\" + AZIP$ + Chr$(13), True
   '
   '
   SendKeys "%A", True
   SendKeys TRIM$(TTXX$) + Chr$(13), True
   '
   AppActivate "WinZip - " + AZIPSH$ + ".ZIP"
   AppActivate "WinZip - " + AZIPSH$ + ".ZIP"
   '
16 On Error Resume Next
   AppActivate "WINZIP"
   If Err Then
       Resume 20
     Else
       SendKeys "%{F4}", True
       GoTo 16
   End If
   '
20 On Error GoTo 0
   '
99 Command2.Enabled = True
   Hide
End Sub

Private Sub Command3_Click()
   '
   Command3.Enabled = False
   CommonDialog1.CancelError = True
   On Error GoTo 99
   '
10 CommonDialog1.DialogTitle = "Crear Archivo de Compresión"
       ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
   CommonDialog1.FILTER = "Zip-Files (*.zip)|*.zip|"
   CommonDialog1.FilterIndex = 1
   CommonDialog1.ShowOpen
   '
   ' Presentar el nombre del archivo seleccionado
   TTXX$ = UCase$(CommonDialog1.FileName)
   PPXX% = InStr(TTXX$, ".")
   If PPXX% > 0 Then
     If Mid$(TTXX$, PPXX% + 1, 3) <> "ZIP" Then
       Call MENSERR(24, "Archivo no Válido - Debe ser '*.ZIP'")
       GoTo 10
     End If
     TTXX$ = Left$(TTXX$, PPXX% - 1)
   End If
   Text2 = TTXX$
   '
98 Resume 99
99 On Error GoTo 0
   Command3.Enabled = True
   '
End Sub

Private Sub Command4_Click()
   '
   Command4.Enabled = False
   For U& = 1 To List1.ListCount
     If List1.Selected(U& - 1) = True Then
       List1.RemoveItem (U& - 1)
       Exit For
     End If
   Next U&
   List1.Refresh
   DoEvents
   Command4.Enabled = True
   '
End Sub

Private Sub Command5_Click()
   '
   Command5.Enabled = False
   For KU% = 0 To 32767
     If EXISTE%(Label11 + "\Documento" + TRIM$(Str$(KU%)) + ".doc") < 1 Then GoTo 10
   Next KU%
   MsgBox "Imposible Abrir en Esta Carpeta"
   GoTo 99
   '
10 Label1 = "Documento" + TRIM$(Str$(KU%)) + ".doc"
   On Error Resume Next
   FileCopy LUDAT$ + "HOJENBLA.DOC", Label11 + "\Documento" + TRIM$(Str$(KU%)) + ".doc"
   If Err > 0 Then
     MsgBox "Error (" + TRIM$(Str$(Err)) + ") al Crear Nuevo Documento."
     On Error GoTo 0
     GoTo 99
   End If
   On Error GoTo 0
   '
   NUEVODOC% = 1
   DOCUORIGI$ = UCase$(TRIM$(Label11 + "\Documento" + TRIM$(Str$(KU%)) + ".doc"))
MsgBox "Rutina Muestradoc Deshabilitada"
   'Call MUESTRADOC
   'MsgBox DOCUORIGI$ + "*" & "   " & UCase$(TRIM$(Label11 + "\Documento" + TRIM$(Str$(KU%)) + ".doc")) + "*"
   If DOCUORIGI$ <> UCase$(TRIM$(Label11 + "\" + Label1)) Then
     On Error Resume Next
     Kill DOCUORIGI$
     On Error GoTo 0
   End If
   NUEVODOC% = 0
   Call CONECTADOC
   '
99 Command5.Enabled = True
Exit Sub
    Command5.Enabled = False
    'narchi$ = InputBox$("Archivo")
    On Error Resume Next
    WORDDOCS.OLE1.Class = "excelsheet"
    WORDDOCS.OLE1.SourceDoc = Label11 + "\" + Label1 '"F:\SOLO LEER\ARCHIVOSCIME-SOI (ACS)\Plantillas Word y Excel\Hojas FLEXOFT"
    WORDDOCS.OLE1.Action = 1
    WORDDOCS.OLE1.Action = 7
    WORDDOCS.VScroll1.Min = 0
    WORDDOCS.VScroll1.Value = 0
    ABARRA = WORDDOCS.OLE1.Height - WORDDOCS.Picture1.Height
    If ABARRA < 1 Then ABARRA = 1
    WORDDOCS.VScroll1.Max = ABARRA
    WORDDOCS.VScroll1.SmallChange = ABARRA / 100
    WORDDOCS.VScroll1.LargeChange = ABARRA / 10
    WORDDOCS.OLE1.Top = 0
    'WORDDOCS.OLE1.DoVerb (-1)
    'SendKeys "%{F10}", True  ' esto para documentos en word
    'SendKeys Chr$(13) + "%{F10}", True ' esto para mails
    'SendKeys Chr$(13) & "%V" & (DOWN8) & "P", True ' ESTO PARA EXCEL
    WORDDOCS.Show 1
    On Error GoTo 0
    Command5.Enabled = True
    '
End Sub

Private Sub Command6_Click()
   Hide
End Sub

Private Sub Dir1_Change()
   Label11.Caption = Dir1.Path
   File1.Path = Dir1.Path
   Label1 = ""
   Call Text1_Change
End Sub

Private Sub Drive1_Change()
   On Error GoTo 20
   Dir1.Path = Drive1.Drive
   GoTo 99
   '
20 Resume 21
21 On Error GoTo 0
   Call MENSERR(24, "Unidad no Disponible.")
   Drive1.Drive = Left$(CurDir$, 2)
   '
99 On Error GoTo 0
End Sub

Private Sub List3_Click()
  Label1 = TRIM$(Left$(List3.TEXT, 32))
End Sub

Private Sub List3_DblClick()
   Call CONECTADOC
End Sub

Private Sub Form_Load()
    '
    Screen.MousePointer = 11
    EPAC$ = EMPREAC$
    If EPAC$ <> "" Then
      Caption = Caption + " - " + EMPREAC$
    End If
    '
    On Error Resume Next
    Drive1.Drive = "F:"
    If Err Then
      On Error GoTo 0
      Drive1.Drive = Left$(App.Path, 2)
    End If
    On Error GoTo 0
    Dir1.Path = "\FLEXOFT\NOVEDAD"
    Label11.Caption = Dir1.Path
    File1.Path = Dir1.Path
    Text1 = "*.EXE"
    Screen.MousePointer = 1
    '
End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Hide
End Sub


Private Sub Label11_Change()
  If Right$(Label11, 1) = "\" Then
    Label11 = Left$(Label11, Len(Label11) - 1)
  End If
End Sub

Private Sub List1_DblClick()
   '
   ILISTA& = List1.ListIndex
   List1.RemoveItem ILISTA&
   '
End Sub

Sub CONECTADOC()
   '
   If TRIM$(Label1) <> "" Then
     XX$ = TRIM$(Label11) + "\" + TRIM$(Label1)
     For U& = 1 To List1.ListCount
       If List1.List(U& - 1) = XX$ Then Exit Sub
     Next U&
'AA1$ = LOADFILE$(XX$)
'AA2$ = String$(Len(AA1$), 0)
'For KU& = 1 To Len(AA1$)
'  CARA% = 1 + Asc(Mid$(AA1$, KU&, 1))
'  If CARA% > 255 Then CARA% = CARA% - 256
'  Mid$(AA2$, KU&, 1) = Chr$(CARA%)
'Next KU&
'PPXX% = InStr(XX$, ".")
'YY$ = Left$(XX$, PPXX%) + "FXC"
'Call SAVEFILE(YY$, AA2$)
     List1.AddItem XX$
     List1.Refresh
     DoEvents
   End If
   '
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 2 Then
     If Y > 0 Then
       YACT% = Int(5 * Y / List1.Height)
       On Error Resume Next
       List1.ListIndex = List1.TopIndex + YACT%
       TTXX$ = List1.TEXT + Chr$(13) + Chr$(10)
       TTXX$ = TTXX$ + "Actualizado: " + Format$(FileDateTime(List1.TEXT), "dd/mm/yyyy - hh:mm") + " hs" + Chr$(13) + Chr$(10)
       TTXX$ = TTXX$ + "Tamaño: " + TRIM$(Str$(FileLen(List1.TEXT))) + " bytes."
       If Err = 0 Then
         MsgBox TTXX$
       End If
       On Error GoTo 0
     End If
   End If
End Sub

Private Sub Option1_Click()
   Text1 = "*.EXE"
End Sub

Private Sub Option2_Click()
   Text1 = "*.PRF"
End Sub

Private Sub Option3_Click()
   Text1 = "*.RFS"
End Sub

Private Sub Option4_Click()
   Text1 = "*.FRM"
End Sub

Private Sub Text1_Change()
   File1.Pattern = Text1
   List2.Clear
   For KU& = 1 To File1.ListCount
     FIDAT$ = Format$(CDbl(FileDateTime(Label11 + "\" + File1.List(KU& - 1))), "00000#.#####   ")
     While Len(FIDAT$) < 15: FIDAT$ = FIDAT$ + " ": Wend
     List2.AddItem FIDAT$ + File1.List(KU& - 1)
   Next KU&
   '
   List3.Clear
   For KU& = List2.ListCount To 1 Step -1
     YY$ = List2.List(KU& - 1)
     IZQUI$ = REPLACAR$(Left$(YY$, 14), ",", ".")
     XX$ = Left$(Mid$(YY$, 16, 32) + Space$(32), 32) + Format$(CDate(Val(IZQUI$)), "dd/mm/yy  hh:mm:ss")
     List3.AddItem XX$
   Next KU&
End Sub

