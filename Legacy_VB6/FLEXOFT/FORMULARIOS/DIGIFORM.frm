VERSION 5.00
Begin VB.Form DIGIFORM 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generador de Formularios de Impresion"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   11910
   Begin VB.CommandButton Command15 
      Caption         =   "Command15"
      Height          =   255
      Left            =   2520
      TabIndex        =   40
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5640
      Left            =   1260
      TabIndex        =   37
      Top             =   1830
      Visible         =   0   'False
      Width           =   8415
   End
   Begin VB.PictureBox Picture3 
      Height          =   435
      Left            =   1260
      ScaleHeight     =   375
      ScaleWidth      =   8355
      TabIndex        =   38
      Top             =   1575
      Visible         =   0   'False
      Width           =   8415
      Begin VB.Label Label1 
         Caption         =   " Numero  Descripción del Documento"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   30
         TabIndex        =   39
         Top             =   30
         Width           =   8205
      End
   End
   Begin VB.CommandButton Command14 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   10920
      Picture         =   "DIGIFORM.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Buscar Documento por Texto Embebido"
      Top             =   6250
      Width           =   855
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Command12"
      Height          =   225
      Left            =   4725
      TabIndex        =   35
      Top             =   0
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   10920
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   33
      Top             =   8085
      Visible         =   0   'False
      Width           =   855
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   -105
         TabIndex        =   34
         Top             =   0
         Width           =   12000
      End
   End
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
      Height          =   400
      Left            =   10920
      Picture         =   "DIGIFORM.frx":1366
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Descarga de Documentos Antiguos / Control y Recuperación de Datos"
      Top             =   7600
      Width           =   855
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Pantalla"
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
      Left            =   10920
      Picture         =   "DIGIFORM.frx":14B0
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Pantalla Completa"
      Top             =   3110
      Width           =   870
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   10920
      Picture         =   "DIGIFORM.frx":17BA
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Imprime Muestra del Formulario Activo"
      Top             =   2420
      UseMaskColor    =   -1  'True
      Width           =   870
   End
   Begin VB.CommandButton Command3 
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   10920
      Picture         =   "DIGIFORM.frx":1E24
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Genera Nuevo Formulario de Impresión"
      Top             =   1745
      Width           =   870
   End
   Begin VB.CommandButton Command99 
      Caption         =   "Recupera"
      Height          =   225
      Left            =   10920
      TabIndex        =   31
      Top             =   8085
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command13 
      Caption         =   "BOTON DE BACKUP"
      Height          =   540
      Left            =   10920
      TabIndex        =   30
      Top             =   -120
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   4545
      Left            =   105
      ScaleHeight     =   4485
      ScaleWidth      =   5700
      TabIndex        =   26
      Top             =   1470
      Visible         =   0   'False
      Width           =   5760
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Logo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   9030
      Picture         =   "DIGIFORM.frx":212E
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Archivo Imagen Isologo"
      Top             =   210
      Width           =   870
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   6960
      LargeChange     =   100
      Left            =   10500
      TabIndex        =   22
      Top             =   1105
      Width           =   225
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   225
      LargeChange     =   100
      Left            =   105
      TabIndex        =   21
      Top             =   8090
      Width           =   10380
   End
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
      Height          =   450
      Left            =   10920
      Picture         =   "DIGIFORM.frx":2438
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Ver Documentos Emitidos"
      Top             =   6700
      Width           =   855
   End
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
      Height          =   450
      Left            =   10920
      Picture         =   "DIGIFORM.frx":2742
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Re-Imprimir Documentos Emitidos"
      Top             =   7150
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Design"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   10920
      TabIndex        =   11
      Top             =   3900
      Width           =   855
      Begin VB.CommandButton BOTIOBJ 
         Caption         =   "Image"
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
         Index           =   5
         Left            =   60
         TabIndex        =   18
         Top             =   1470
         Width           =   730
      End
      Begin VB.CommandButton BOTIOBJ 
         Caption         =   "Parms"
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
         Index           =   4
         Left            =   60
         TabIndex        =   16
         Top             =   1820
         Width           =   730
      End
      Begin VB.CommandButton BOTIOBJ 
         Caption         =   "Box"
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
         Index           =   3
         Left            =   60
         TabIndex        =   15
         Top             =   1170
         Width           =   730
      End
      Begin VB.CommandButton BOTIOBJ 
         Caption         =   "Line"
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
         Index           =   2
         Left            =   60
         TabIndex        =   14
         Top             =   870
         Width           =   730
      End
      Begin VB.CommandButton BOTIOBJ 
         Caption         =   "Label"
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
         Index           =   1
         Left            =   60
         TabIndex        =   13
         Top             =   570
         Width           =   730
      End
      Begin VB.CommandButton BOTIOBJ 
         Caption         =   "Gral"
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
         Index           =   0
         Left            =   60
         TabIndex        =   12
         Top             =   230
         Width           =   730
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   9975
      Picture         =   "DIGIFORM.frx":2DAC
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Guarda Cambios en Formulario Activo"
      Top             =   210
      Width           =   870
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Formulario Activo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   105
      TabIndex        =   3
      Top             =   105
      Width           =   8730
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3465
         MultiLine       =   -1  'True
         TabIndex        =   7
         Text            =   "DIGIFORM.frx":30B6
         Top             =   370
         Width           =   5055
      End
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
         Height          =   360
         Left            =   3150
         TabIndex        =   6
         Top             =   370
         Width           =   170
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1260
         TabIndex        =   5
         Text            =   " "
         Top             =   370
         Width           =   1905
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label3"
         Height          =   255
         Left            =   4680
         TabIndex        =   41
         Top             =   0
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
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
         Left            =   525
         TabIndex        =   4
         Top             =   525
         Width           =   855
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   10920
      Picture         =   "DIGIFORM.frx":30BA
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Ayuda Flexoft en Línea"
      Top             =   1050
      Width           =   870
   End
   Begin VB.CommandButton Command2 
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
      Height          =   800
      Left            =   10920
      Picture         =   "DIGIFORM.frx":33C4
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Termina - Cierra la Aplicación"
      Top             =   210
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6980
      Left            =   105
      TabIndex        =   23
      Top             =   1105
      Width           =   10400
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   6960
         Left            =   0
         ScaleHeight     =   6900
         ScaleWidth      =   10320
         TabIndex        =   24
         Top             =   0
         Width           =   10380
      End
      Begin VB.Line Line3 
         X1              =   10395
         X2              =   0
         Y1              =   6960
         Y2              =   6960
      End
      Begin VB.Line Line1 
         X1              =   10380
         X2              =   10380
         Y1              =   0
         Y2              =   6930
      End
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   4725
      Pattern         =   "*.FRM"
      TabIndex        =   17
      Top             =   6930
      Visible         =   0   'False
      Width           =   2745
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   1365
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   6930
      Visible         =   0   'False
      Width           =   2745
   End
   Begin VB.Label TEPRUORI 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   105
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Label TEPRURED 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   0
      TabIndex        =   27
      Top             =   315
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Line Line2 
      X1              =   10725
      X2              =   10725
      Y1              =   1105
      Y2              =   8050
   End
End
Attribute VB_Name = "DIGIFORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim XMOUSE, YMOUSE
'
Private Sub BOTIOBJ_Click(Index As Integer)
    If TRIM$(CODFOR$) = "" Then
       Call MENSERR(24, "Falta Seleccionar Formulario")
       Text1.SetFocus
       Exit Sub
    End If
    '
    DESFOR$ = TRIM$(Text2)
    Select Case Index
      Case 0
        Call DEFLAYOUTW
      Case 1
        Call DEFLEYENW
      Case 2
        Call DEFLINEAW
      Case 3
        Call DEFCAJASW
      Case 4
        Call DEFPARAMW
      Case 5
        Call DEFIMAGEW
    End Select
    Call SHOWFORM
End Sub

Private Sub Command1_Click()
    '
    Call SELECHO("INDIFRM")
    CODFOR$ = VALACT1$("INDIFRM")
    If TRIM$(CODFOR$) = "" Then
        Text1.SetFocus
      Else
        Text1.TEXT = CODFOR$
        Text2.TEXT = ECOARCH$("INDIFRM", CODFOR$)
        Call SHOWFORM
    End If
    '
End Sub

Private Sub Command10_Click()
   Command10.Enabled = False
   FORILOGO.Show 1
   Command10.Enabled = True
End Sub

Private Sub Command11_Click()
   '
   Command11.Enabled = False
   OPPX% = ALTERNA%("Descarga Documentos Viejos\Control y Recuperación de Datos\Control de Documentos")
   If OPPX% = 1 Then
        Call DESCADOC
      ElseIf OPPX% = 2 Then
        Call CHECKDOC
      ElseIf OPPX% = 3 Then
        Call VERIDOC
   End If
   '
99 Screen.MousePointer = 1
   Command11.Enabled = True
   MARCOBARRA.Visible = False
   End Sub

Sub DESCADOC()
   '
   Call COMUNI("No Disponible en la Presente Version.")
   Exit Sub
   '
   If DERACCE%("DESCADOC", "Descarga de Base de Datos de Documentos") >= 2 Then
     If COMALTER%("Realmente Quiere Descargar la Base de Datos\de Documentos Impresos ?\\No, Conservar\Si, Descargar.") = 2 Then
       MARCOBARRA.Visible = True
       Call BLOQARCH("PDOCSPL")
       On Error Resume Next
       Name LUDAT$ + "PDOCSPL.DAT" As LUDAT$ + "PDOCSPL0.DAT"
       If Err Then
         On Error GoTo 0
         Call MENSERR(24, "Proceso Imposible.\   \Cierre FLEXOFT en TODAS las\Estaciones de Trabajo, y\Repita la Petición.")
         Call CIERRARCH("MENSAJES")
         GoTo 99
       End If
       On Error GoTo 0
       '
       Screen.MousePointer = 11
       Call COPYSTRU("PDOCSPL", "PDOCSPL0")
       '
       FIN& = ULTREG&("PDOCLST")
       For U& = FIN& To 1 Step -1
         XX$ = REGLEIDO$("PDOCLST", U&)
         UDOC& = CVS(Left$(XX$, 4))
         If UDOC& > 0 Then GoTo 20
       Next U&
       UDOC& = 1
       '
20     FINDESCA& = UDOC& / 2   ' descarga la mitad del archivo
       Dim TD0 As String * 512
       NX0% = FILEOPEN%(LUDAT$ + "PDOCSPL.DT0", 0, 512)
       URE0& = LOF(NX0%) / 512
       NXXX% = FILENBR%("PDOCSPL0")
       LOFI = LOF(NXXX%)
       FIN& = LOFI / 512
       If LOFI < 0 Then FIN& = 4194304
       '
       Call CIERRARCH("PDOCSPL0")

       NX1% = FILEOPEN%(LUDAT$ + "PDOCSPL.DAT", 0, 512)
       NXXX% = FILEOPEN%(LUDAT$ + "PDOCSPL0.DAT", 0, 512)
       URE1& = 0
       '
       Dim XX00 As String * 512
       For U& = 1 To FIN&
         Call TRACE(20, U&, FIN&)
         Get #NXXX%, U&, XX00$         ' XX$ = REGLEIDO$("PDOCSPL0", U&)
         XX$ = XX00$
         TD0$ = XX$
         If CVS(Left$(XX$, 4)) <= FINDESCA& Then
              URE0& = URE0& + 1
              Put #NX0%, URE0&, TD0$
            Else
              URE1& = URE1& + 1
              Put #NX1%, URE1&, TD0$
         End If
       Next U&
       '
       Call CIERRARCH("PDOCSPL0")
       Close #NX0%
       Close #NX1%
       Close #NXXX%
       '
       Screen.MousePointer = 1
       On Error Resume Next
       Kill LUDAT$ + "PDOCSPL0.DAT"
       If Err Then
           On Error GoTo 0
           Call COMUNI("No se Pudo Completar el Proceso - Consulte.")
         Else
           '\\\OT-07-0111-RG-15-03-07/// 'MENSAJE QUE MUESTRA CUANDO FINALIZA EL PROCESO
           On Error GoTo 0
           Call COMUNI("Fin de Proceso de Descarga:\   \Operación Exitosa\")
           '\\\OT-07-0111-RG-FIN///
       End If
       On Error GoTo 0
       '
     End If
   End If
   '
99 End Sub

Sub CHECKDOC()
   '
   Dim RQX As String * 512
   Dim TD0 As String * 512
   '
   If DERACCE%("CHECKDOC", "Control y Recuperación de B.D. Documentos") >= 2 Then
     If COMALTER%("Realmente Quiere Recuperar la Base de Datos\de Documentos Impresos ?\\No, Cancelar\Si, Recuperar.") = 2 Then
       MARCOBARRA.Visible = True
       'OPX% = COMALTER%("Para este Proceso debe Cerrarse FLEXOFT\en TODAS las Estaciones de Trabajo.\   \Pulse 'Continuar' para Iniciar el Proceso\\Continuar\Cancelar")
       'If OPX% <> 1 Then Exit Sub
       '
       EPAC$ = UCase$(EMPREAC$)
       If InStr(EPAC$, "DOSIVAC") > 0 Then
          Call CHECKDOSI
          Exit Sub
       End If
       '
       Label3 = ""
       Label3.Visible = True
       TINI = Timer
       '
       ARECU$ = "RDOCSPL.DT0"
       NQX% = FILEOPEN%(LUDAT$ + ARECU$, 0, 512)
       JJ& = 0
       '
       U& = 1: PUSOAVISO% = 0
       Do While U& <= ULTREG&("PDOCLST")
         FIN& = ULTREG&("PDOCLST")
         SEGNEC = ((Timer - TINI) * (FIN& - U&) / U&)
         DIANEC = SEGNEC / (24# * 3600)
         Label3 = "Procesados: " + CStr(U&) + " - Total: " + CStr(FIN&) + " - Falta: " + Format(DIANEC, "HH:MM:SS")
         Label3.Refresh: DoEvents
         XX$ = REGLEIDO$("PDOCLST", U&)
         DOCUNU& = CVS(Left$(XX$, 4))
         GoSub RECUDOCU
         '
         If U& > FIN& - 2000 Then
           If PUSOAVISO% < 1 Then
             Call COMUNI("Falta Poco Para Terminar.\   \Todos los Usuarios Deben Salir del FLEXOFT\y Cerrar el FLEXOFT en TODAS las Estaciones.\   \Luego 'Aceptar' para Continuar Proceso.")
             PUSOAVISO% = 1
           End If
         End If
         '
       U& = U& + 1
       Loop
       '
       Call COMUNI("FIN DE PROCESO\   \Debe ahora Renombrar los Archivos 'RDOCSPL.DT_' como 'PDOCSPL.DT_'")
       '
     End If
   End If
   '
Exit Sub
'
'
RECUDOCU:
   ABUS$ = "PDOCSPL.DT0"
5  NX0% = FILEOPEN%(LUDAT$ + ABUS$, 0, 512)
   '
   URE& = Abs(LOF(NX0%) / 512)   ' PARA PREVENIR DESBORDAMIENTO DOSIVAC
   LI& = 0: LS& = URE&
6  E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
   L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
   '
   Get #NX0%, L&, TD0$
   DCC& = CVS(Left$(TD0$, 4))
   If DCC& < DOCUNU& Then LI& = L&: GoTo 6
   If DCC& > DOCUNU& Then LS& = L&: GoTo 6
   GoTo 8
   '
7  Close #NX0%: NX0% = 0
   If ABUS$ = "PDOCSPL.DT0" Then
        ABUS$ = "PDOCSPL.DT1"
        GoTo 5
      ElseIf ABUS$ = "PDOCSPL.DT1" Then
        ABUS$ = "PDOCSPL.NEW"
        GoTo 5
      ElseIf ABUS$ = "PDOCSPL.NEW" Then
        ABUS$ = "PDOCSPL.DAT"
        GoTo 5
   End If
   'Call MENSERR(24, "Documento " & DOCUNU& & " no Registrado")
   GoTo 99
   '
8  TPSP$ = ""
   Screen.MousePointer = 11
   While L& > 1
     Get #NX0%, L&, TD0$
     DCC& = CVS(Left$(TD0$, 4))
     If DCC& < DOCUNU& Then GoTo 9
     L& = L& - 1
   Wend
   '
9  For KKL& = L& To URE&
     Get #NX0%, KKL&, TD0$
     TBUF$ = TD0$
     If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 99
     If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
       JJ& = JJ& + 1
       RQX$ = TD0$
       Put #NQX%, JJ&, RQX$
     End If
   Next KKL&
   '
99 If LOF(NQX%) > 2000000000 Then
     If ARECU$ = "RDOCSPL.DT0" Then
       Close #NQX%
       ARECU$ = "RDOCSPL.DAT"
       NQX% = FILEOPEN%(LUDAT$ + ARECU$, 0, 512)
       JJ& = 0
     End If
   End If
   '
   If NX0% > 0 Then
      Close #NX0%
      NX0% = 0
   End If
Return
'
'
End Sub

Sub VERIDOC()
    '
    Call COMUNI("No Disponible en la Presente Version.")
    Exit Sub
    '
    Dim ISDOC%(), DESRE&(), HASRE&()
    '
    MARCOBARRA.Visible = True
    FIN& = ULTREG&("PDOCLST")
    For U& = FIN& To 1 Step -1
       XX$ = REGLEIDO$("PDOCLST", U&)
       UDOC& = CVS(Left$(XX$, 4))
       If UDOC& > 0 Then GoTo 20
    Next U&
    UDOC& = 1
    '
20  UDOC& = UDOC& + 256
    ReDim ISDOC%(UDOC&), DESRE&(UDOC&), HASRE&(UDOC&)
    '
    PRIDOC& = 0: ULTDOC& = 0
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       XX$ = REGLEIDO$("PDOCLST", U&)
       NDOC& = CVS(Left$(XX$, 4))
' VERIFICAR AQUI ORDEN CORRELATIVO
       If NDOC& > 0 Then
         If PRIDOC& = 0 Then PRIDOC& = NDOC&
         If NDOC& <= UDOC& Then
           ISDOC%(NDOC&) = 1
           If NDOC& > ULTDOC& Then ULTDOC& = NDOC&
           If NDOC& <= DOCANT& Then
              YY$ = TRIM$(Mid$(XX$, 5, 60))
              MsgBox "Registro " & U& & " - " & YY$ & " - Desorden en Indice."
           End If
           DOCANT& = NDOC&
           GoTo 29
         End If
       End If
29   Next U&
     '
     'NX0% = FILEOPEN%(LUDAT$ + "PDOCSPL.DT0", 0, 512)
     'URE0& = LOF(NX0%) / 512
     'FIN& = ULTREG&("PDOCSPL0")
     'NX1% = FILEOPEN%(LUDAT$ + "PDOCSPL.DAT", 0, 512)
     'URE1& = 0
     '
     FIN& = ULTREG&("PDOCSPL")
     DOCANT& = 0
     For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       XX$ = REGLEIDO$("PDOCSPL", U&)
       NDOC& = CVS(Left$(XX$, 4))
       If NDOC& < DOCANT& Then
           YY$ = TRIM$(Mid$(XX$, 5, 60))
           MsgBox "Registro " & U& & " de PDOCSPL - Desorden en Detalle. (" & DOCANT& & " / " & NDOC& & ")."
       End If
       DOCANT& = NDOC&
       If NDOC& > 0 Then
         If NDOC& <= UDOC& Then
           If DESRE&(NDOC&) < 1 Then DESRE&(NDOC&) = U&
           If HASRE&(NDOC&) < U& Then HASRE&(NDOC&) = U&
         End If
       End If
     Next U&
     '
     For U& = 1 To UDOC&
        If ISDOC%(U&) = 1 Then
           If DESRE&(U&) < 1 Or HASRE&(U&) < DESRE&(U&) Then
              YY$ = TRIM$(Mid$(REGLEIDO$("PDOCLST", U&), 5, 60))
              MsgBox "Registro " & U& & " - " & YY$ & " - Falta Detalle "
           End If
           If DESRE&(U&) > 0 Then
             If DESRE&(U&) <= HASRE&(U& - 1) Then
               MsgBox "Punteros Cruzados"
             End If
           End If
        End If
     Next U&
     '
End Sub

Sub CHECKDOSI()
   '
   Call BLOQARCH("PDOCSPL")
   FIDATI# = FIDATIM#(LUDAT$ + "PDOCSPL.DAT")
   '
   Screen.MousePointer = 11
   Call COPYSTRU("PDOCSPL", "PDOCSPL0")
   On Error Resume Next
   Kill LUDAT$ + "PDOCSPL0.DAT"
   On Error GoTo 0
   '
   If EXISTE%(LUDAT$ + "PDOCSPL0.DAT") > 0 Then
      Call MENSERR(24, "Proceso Imposible. No fue Posible Borrar\Archivo de Transferencia 'PDOCSPL0.DAT'")
      GoTo 99
   End If
   '
   List2.Visible = True
   Picture3.Visible = True
   '
   YY$ = REGLEIDO$("PDOCLST", 1)
   MINI = CVS(Left$(YY$, 4))
   URELI& = ULTREG&("PDOCLST")
   YY$ = REGLEIDO$("PDOCLST", URELI&)
   MAXI = CVS(Left$(YY$, 4))
   '
   FIN& = ULTREG&("PDOCSPL")
   JJ& = 0: NDOCA& = 0: CAREDEL& = 0
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      VALOPRO& = 1 + (U& / FIN&) * (MAXI - MINI + 1)
      XX$ = REGLEIDO$("PDOCSPL", U&)
      VALEIDO = CVS(Left$(XX$, 4))
      If (Abs(VALOPRO& - VALEIDO)) < 10 Or ((Abs(VALOPRO& - VALEIDO)) / VALOPRO&) < 0.1 Then
         JJ& = JJ& + 1
         Call GRAREG("PDOCSPL0", XX$, JJ&)
         NDOC1& = VALEIDO
         If NDOC1& <> NDOCA& Then
           YY$ = FILTERGETRECORD$("PDOCLST", 1, MKS$(NDOC1&))
           ZZ$ = FORMATNUM$(NDOC1&, "F7.0") + "  " + Mid$(YY$, 5, 60)
           If List2.ListCount > 30000 Then
              List2.RemoveItem 0
           End If
           List2.AddItem ZZ$
           List2.Visible = True
           List2.ListIndex = List2.ListCount - 1
           List2.Refresh
           NDOCA& = NDOC1&
         End If
         CAREDEL& = 0
         GoTo 29
      End If
      If CAREDEL& = 0 Then
        List2.AddItem "   **********************  Documentos Perdidos  *************************"
      End If
      CAREDEL& = CAREDEL& + 1
      '
29  Next U&
   '
   Call CIERRARCH("PDOCSPL")
   Call CIERRARCH("PDOCSPL0")
   '
   If ULTREG&("PDOCSPL") <> FIN& Then
      Call MENSERR(24, "Imposible Continuar. Usuarios Conectados\Activos Durante el Proceso.")
      GoTo 99
   End If
   '
   Close
   Reset
   Call BLOQARCH("PDOCSPL")
   On Error Resume Next
   Name LUDAT$ + "PDOCSPL.DAT" As LUDAT$ + "PDOCSPL.BAD"
   If Err Then
      On Error GoTo 0
      Call MENSERR(24, "Proceso Interrumpido. No fue Posible\Re-Nombrar 'PDOCSPL.DAT' como 'PDOCSPL.BAD'.\   \Repita el Proceso.")
      GoTo 99
   End If
   On Error GoTo 0
   '
   On Error Resume Next
   Name LUDAT$ + "PDOCSPL0.DAT" As LUDAT$ + "PDOCSPL.DAT"
   If Err Then
      On Error GoTo 0
      Call MENSERR(24, "Proceso Interrumpido. No fue Posible\Re-Nombrar 'PDOCSPL0.DAT' como 'PDOCSPL.DAT'.\   \Consulte Urgente a Soporte FLEXOFT.")
      GoTo 99
   End If
   On Error GoTo 0
   '
   Call COMUNI("Fin de Proceso de Recuperación:\   \Operación Exitosa\")
99 On Error GoTo 0
   '
End Sub
'
Private Sub Command12_Click()
   LUACC$ = "F:\CLIENTES\DOSIVAC\DATOS\BDACCESS\"
   Set BADAFORM = OpenDatabase(LUACC$ + "BADAFORM.MDB")
   Set PDOCLIST = BADAFORM.OpenRecordset("PDOCLIST", dbOpenDynaset)
   '
   For KKI% = 0 To 5
     EDIFORMU.Picture1(KKI%).Cls
     EDIFORMU.Picture1(KKI%).Top = 0
     EDIFORMU.Picture1(KKI%).Refresh
     EDIFORMU.Picture1(KKI%).Height = 7070
     EDIFORMU.Picture1(KKI%).Width = 11700
     EDIFORMU.Picture1(KKI%).DrawWidth = 3
     EDIFORMU.Picture1(KKI%).DrawStyle = 0
     EDIFORMU.Picture1(KKI%).FillStyle = 1
   Next KKI%
   ALTUPAGINA = 7070
   TOPEPAGINA = 0
   PAGINACTIVA% = 0
   CAPAGINAS% = 0
   '
5  Call SELECHO("PDOCLST")
   DOCUNU& = Val(VALACT1$("PDOCLST"))
   If DOCUNU& < 1 Then
      Exit Sub
   End If
   '
   Screen.MousePointer = 11
   Set RESEL = BADAFORM.OpenRecordset("SELECT * FROM PDOCLIST WHERE NUMERO = " & DOCUNU, dbOpenSnapshot)
   RESEL.MoveFirst
   Screen.MousePointer = 1
   '
95 TPSP$ = RESEL!BINARIO
   '
10 PPXX% = InStr(TPSP$, Chr$(255))
   If PPXX% > 0 Then
      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
      TPSP$ = Mid$(TPSP$, PPXX% + 1)
      '
      If Len(COPRIAT$) > 0 Then
        COPRI% = Asc(Left$(COPRIAT$, 1))
        ATRIP$ = Mid$(COPRIAT$, 2)
        Call SHOWCOMMAND(COPRI%, ATRIP$)
      End If
      GoTo 10
      '
   End If
   '
   Screen.MousePointer = 1
   PAGINACTIVA% = 0
   ANCHPAGINA = EDIFORMU.Picture1(PAGINACTIVA%).Width
   EDIFORMU.Picture1(PAGINACTIVA%).Refresh
   EDIFORMU.VScroll1.Min = 0
   EDIFORMU.VScroll1.Value = 0
   EDIFORMU.VScroll1.Max = 1
   If ALTUPAGINA > EDIFORMU.Frame1.Height Then
      EDIFORMU.VScroll1.Max = (ALTUPAGINA - EDIFORMU.Frame1.Height) / 100
   End If
   EDIFORMU.HScroll1.Min = 0
   EDIFORMU.HScroll1.Value = 0
   EDIFORMU.HScroll1.Max = 1
   If ANCHPAGINA > EDIFORMU.Frame1.Width Then
      EDIFORMU.HScroll1.Max = (ANCHPAGINA - EDIFORMU.Frame1.Width) / 100
   End If
   '
   Call CENTRAFORM(EDIFORMU)
   EDIFORMU.Show
   '
   
End Sub

Private Sub Command13_Click()
   DIREBAK$ = "C:\FLEXOFT\BACKUP"
   On Error Resume Next
   MkDir DIREBAK$
   On Error GoTo 0
   '
   HOII% = HOY(HOS$)
   ADIAR$ = "BK" + Mid$(HOS$, 7, 2) + Mid$(HOS$, 4, 2) + Mid$(HOS$, 1, 2)
   '
   ARESPA$ = LUDAT$ + "*.DAT " + LUDAT$ + "*.OLD " + LUDAT$ + "*.FRM " + LUDAT$ + "*.X*"
   CMDZ$ = "ARJ A " + DIREBAK$ + "\" + ADIAR$ + " " + ARESPA$
   'MsgBox CMDZ$
   'Call DOSCOMMAND(CMDZ$)
   Call SAVEFILE("C:\FLEXOFT\FLEXBAT.BAT", CMDZ$)
   Shell "C:\FLEXOFT\DOSCOMM.PIF", vbNormalFocus

End Sub

Private Sub Command14_Click()
    '
    Command14.Enabled = False
    '
10  TEBUS$ = InputBox$("Busqueda por Texto", "Introduzca Argumento de Búsqueda", TEBUS$)
    If TEBUS$ = "" Then GoTo 99
    '
    Screen.MousePointer = 11
    MARCOBARRA.Visible = True
    Call COPYSTRU("PDOCLST", "PDOCLIS")
    URE& = ULTREG&("PDOCSPL")
    Call SETULTREG("PDOCLIS", 0)
    Z& = 0
    For I& = 1 To ULTREG&("PDOCSPL")
      Call TRACE(20, I&, URE&)
      RELE$ = REGLEIDO$("PDOCSPL", I&)
      X$ = Mid$(RELE$, 5)
      
      If InStr(1, X$, TEBUS$) > 0 Then
         DOCUNU& = CVS(Left$(RELE$, 4))
         '
         RECORD$ = FILTERGETRECORD$("PDOCLST", 1, MKS$(DOCUNU&))
         '
         Z& = Z& + 1
         Call GRAREG("PDOCLIS", RECORD$, Z&)
         '
         Call CIERRARCH("PDOCLIS")
      End If
    Next I&
    Call SETULTREG("PDOCLIS", Z&)
    MARCOBARRA.Visible = False
    '
    If Z& < 1 Then
       Call MENSERR(24, "No se Encontraron Documentos\para el Argumento de Búsqueda.")
       GoTo 10
    End If
    '
25  Screen.MousePointer = 1
    Call SELECHO("PDOCLIS/Búsqueda en Documentos por Texto: '" + TEBUS$ + "'")
    DOCUNU& = Val(VALACT1$("PDOCLIS"))
    If DOCUNU& > 0 Then
       Call MUESTRA_DOCU(DOCUNU&)
       GoTo 25
    End If
    GoTo 10
    '
99  Command14.Enabled = True
    '
End Sub

Private Sub Command15_Click()
    '
    On Error Resume Next
    Name LUDAT$ + "PDOCSPL.DT0" As LUDAT$ + "PDOCSPL0.DAT"
    Name LUDAT$ + "PDOCSPL.DT1" As LUDAT$ + "PDOCSPL1.DAT"
    On Error GoTo 0
    '
    Call COPYSTRU("PDOCSPL", "PDOCSPL0")
    Call COPYSTRU("PDOCSPL", "PDOCSPL1")
    '
    Dim RQX As String * 512
    NQX% = FILEOPEN%(LUDAT$ + "PDOCSPL.DT0", 0, 512)
    JJ& = 0
    '
    FIN& = ULTREG&("PDOCLST")
    For U& = 1 To FIN&
      Label3 = U&
      Label3.Refresh: DoEvents
      XX$ = REGLEIDO$("PDOCLST", U&)
      DOCUNU& = CVS(Left$(XX$, 4))
      GoSub RECUDOCU
    Next U&
    '
    Call COMUNI("FIN DE PROCESO")
    '
Exit Sub
'
RECUDOCU:
   ABUS$ = "PDOCSPL0"
   Dim TD0 As String * 512
5  NX0% = FILEOPEN%(LUDAT$ + ABUS$ + ".DAT", 0, 512)
   '
   URE& = Abs(LOF(NX0%) / 512)   ' PARA PREVENIR DESBORDAMIENTO DOSIVAC
   LI& = 0: LS& = URE&
6  E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
   L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
   '
   Get #NX0%, L&, TD0$
   DCC& = CVS(Left$(TD0$, 4))
   If DCC& < DOCUNU& Then LI& = L&: GoTo 6
   If DCC& > DOCUNU& Then LS& = L&: GoTo 6
   GoTo 8
   '
7  Close #NX0%: NX0% = 0
   If ABUS$ = "PDOCSPL0" Then
      ABUS$ = "PDOCSPL1"
      GoTo 5
   End If
   Call MENSERR(24, "Documento " & DOCUNU& & " no Registrado")
   GoTo 99
   '
8  TPSP$ = ""
   Screen.MousePointer = 11
   While L& > 1
     Get #NX0%, L&, TD0$
     DCC& = CVS(Left$(TD0$, 4))
     If DCC& < DOCUNU& Then GoTo 9
     L& = L& - 1
   Wend
   '
9  For KKL& = L& To URE&
     Get #NX0%, KKL&, TD0$
     TBUF$ = TD0$
     If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 99
     If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
       JJ& = JJ& + 1
       RQX$ = TD0$
       Put #NQX%, JJ&, RQX$
     End If
   Next KKL&
   '
99 If NX0% > 0 Then
      Close #NX0%
      NX0% = 0
   End If
Return
   '
End Sub

Private Sub Command2_Click()
    Call CIERRARCH("*.*")
    Close
    Call FINAL("")
End Sub

Private Sub Command3_Click()
    Call ALTANUEFORM
End Sub

Private Sub Command4_Click()
    '
    If TRIM$(CODFOR$) = "" Then
       Call MENSERR(24, "Falta Seleccionar Formulario")
       Text1.SetFocus
       Exit Sub
    End If
    '
    MARIZW = Val(ATRIFORM$(CODFOR$, "MARIZQMM"))
    MARSUW = Val(ATRIFORM$(CODFOR$, "MARSUPMM"))
    '
    PORIAN = Printer.Orientation  ' retiene orientacion
    On Error Resume Next
    Printer.Orientation = 1       ' anterior
    If Err Then
       Call MENSERR(24, "Imposible Imprimir.\Error de Comunicacion con Impresora.")
       Resume
       On Error GoTo 0
       Text1.SetFocus
       Exit Sub
    End If
    On Error GoTo 0
    '
    PORI$ = TRIM$(UCase$(ATRIFORM$(CODFOR$, "ORIENTA")))
    If PORI$ = "PAISAJE" Or PORI$ = "2" Then
        Printer.Orientation = 2
    End If
    '
    CALIAN = Printer.PrintQuality
    Printer.PrintQuality = (-1) ' DRAFT
    '
    FORMALTER.PSPOOL.Clear
    Call CARFORWIN
    'Printer.EndDoc
    Call SETSPOOL("ENDDOC", "")
    DESCRIDOC$ = "Formulario " + TRIM$(Text1.TEXT) + " - " + TRIM$(Text2.TEXT)
    HOII% = HOY(HOS$)
    DESCRIDOC$ = DESCRIDOC$ + " - " + HOS$
    Call SAVESPOOL(DESCRIDOC$, OKX%)
    '
    On Error Resume Next
    Printer.Orientation = PORIAN
    Printer.PrintQuality = CALIAN
    On Error GoTo 0
    '
End Sub

Private Sub Command5_Click()
    '
    If TRIM$(CODFOR$) = "" Then
       Call MENSERR(24, "Falta Seleccionar Formulario")
       Text1.SetFocus
       Exit Sub
    End If
    '
    Call GRATITUFORM
    '
End Sub

Private Sub Command6_Click()
    Call HELPONLINE("PRIFORMS")
End Sub

Private Sub Command7_Click()
   '
   On Error Resume Next
   Printer.DrawWidth = 3
   Printer.DrawStyle = 0
   Printer.FillStyle = 1
   If Err Then
     Call MENSERR(24, "Imposible Imprimir.\Error de Comunicacion con la Impresora.")
     Resume
     On Error GoTo 0
     Exit Sub
   End If
   On Error GoTo 0
   '
5  Call SELECHO("PDOCLST")
   DOCUNU& = Val(VALACT1$("PDOCLST"))
   If DOCUNU& < 1 Then
      Exit Sub
   End If
   REDOCU$ = REGACT$("PDOCLST")
   '
   TPSP$ = ""
   Screen.MousePointer = 11
   '
   Dim TD0 As String * 512
   NX0% = FILEOPEN%(LUDAT$ + "PDOCSPL.DAT", 0, 512)
   Get #NX0%, 1, TD0$
   If CVS(Left$(TD0$, 4)) > DOCUNU& Then
     Close #NX0%
     NX0% = FILEOPEN%(LUDAT$ + "PDOCSPL.DT0", 0, 512)
   End If
   '
   URE& = Abs(LOF(NX0%) / 512)   ' PARA PREVENIR DESBORDAMIENTO DOSIVAC
   LI& = 0: LS& = URE&
6  E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
   L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
   Get #NX0%, L&, TD0$
   DCC& = CVS(Left$(TD0$, 4))
   If DCC& < DOCUNU& Then LI& = L&: GoTo 6
   If DCC& > DOCUNU& Then LS& = L&: GoTo 6
   GoTo 8
   '
7  Call MENSERR(24, "Documento no Registrado")
   GoTo 5
   '
8  TPSP$ = ""
   Screen.MousePointer = 11
   While L& > 1
     Get #NX0%, L&, TD0$
     DCC& = CVS(Left$(TD0$, 4))
     If DCC& < DOCUNU& Then GoTo 9
     L& = L& - 1
   Wend
   '
9  For KKL& = L& To URE&          'SE CAMBIO EL ARCHIVO DE LECTURA
     Get #NX0%, KKL&, TD0$
     TBUF$ = TD0$
     If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
         TPSP$ = TPSP$ + Mid$(TBUF$, 5)
       ElseIf CVS(Left$(TBUF$, 4)) > DOCUNU& Then
         GoTo 91
     End If
   Next KKL&
   '
91 Close #NX0%
   Screen.MousePointer = 11
   On Error Resume Next
   Printer.PrintQuality = (-1)
   MATRIXPUN% = 0
   If InStr(Printer.DeviceName, "810") > 0 Then MATRIXPUN% = 1
   If InStr(Printer.DeviceName, "286") > 0 Then MATRIXPUN% = 1
   If InStr(UCase$(Printer.DeviceName), "APEX") > 0 Then MATRIXPUN% = 1
   If InStr(UCase$(Printer.DeviceName), "FACTURA") > 0 Then MATRIXPUN% = 1
   If InStr(UCase$(Printer.DeviceName), "REMITO") > 0 Then MATRIXPUN% = 1
   'If InStr(Printer.DriverName, "EPSON9") > 0 Then MATRIXPUN% = 1
   FORMALTER.List1.Clear
   On Error GoTo 0
   '
   Printer.EndDoc        ' RESETEA IMPRESORA
   '
10 PPXX% = InStr(TPSP$, Chr$(255))
   If PPXX% > 0 Then
      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
      TPSP$ = Mid$(TPSP$, PPXX% + 1)
      '
      If Len(COPRIAT$) > 0 Then
        COPRI% = Asc(Left$(COPRIAT$, 1))
        ATRIP$ = Mid$(COPRIAT$, 2)
        Call PRINTCOMMAND(COPRI%, ATRIP$)
      End If
      GoTo 10
      '
   End If
   '
   Printer.EndDoc
   Screen.MousePointer = 1
   GoTo 5
   '
End Sub
'
Private Sub Command8_Click()
   '
   Command8.Enabled = False
   If EXISTE%("FLFORM09.EXE") > 0 Then
      Call CONECRUN("FLFORM09", "Ver Documentos Emitidos")
      GoTo 99
   End If
   '
   For KKI% = 0 To 5
     EDIFORMU.Picture1(KKI%).Cls
     EDIFORMU.Picture1(KKI%).Top = 0
     EDIFORMU.Picture1(KKI%).Refresh
     EDIFORMU.Picture1(KKI%).Height = 7070
     EDIFORMU.Picture1(KKI%).Width = 11700
     EDIFORMU.Picture1(KKI%).DrawWidth = 3
     EDIFORMU.Picture1(KKI%).DrawStyle = 0
     EDIFORMU.Picture1(KKI%).FillStyle = 1
   Next KKI%
   ALTUPAGINA = 7070
   TOPEPAGINA = 0
   PAGINACTIVA% = 0
   CAPAGINAS% = 0
   '
5  Call SELECHO("PDOCLST")
   DOCUNU& = Val(VALACT1$("PDOCLST"))
   If DOCUNU& < 1 Then
      Exit Sub
   End If
   REDOCU$ = REGACT$("PDOCLST")
   ORIPRINTER$ = TRIM$(Mid$(REDOCU$, 129))
   '
AORI$ = "P": If EXISTE%(LUDAT$ + "RDOCSPL.DT0") > 0 Then AORI$ = "R"
   Dim TD0 As String * 512
   NX0% = FILEOPEN%(LUDAT$ + AORI$ + "DOCSPL.DAT", 0, 512)
   Get #NX0%, 1, TD0$
   If CVS(Left$(TD0$, 4)) > DOCUNU& Or LOF(NX0%) < 1 Then
     Close #NX0%
     NX0% = FILEOPEN%(LUDAT$ + AORI$ + "DOCSPL.DT0", 0, 512)
   End If
   '
   URE& = Abs(LOF(NX0%) / 512)   ' PARA PREVENIR DESBORDAMIENTO DOSIVAC
   LI& = 0: LS& = URE&
6  E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
   L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
   '
   Get #NX0%, L&, TD0$
   DCC& = CVS(Left$(TD0$, 4))
'AA1 = Len(TD0$)
   If DCC& < DOCUNU& Then LI& = L&: GoTo 6
   If DCC& > DOCUNU& Then LS& = L&: GoTo 6
   GoTo 8
   '
7  Call MENSERR(24, "Documento no Registrado")
   GoTo 5
   '
8  TPSP$ = ""
   Screen.MousePointer = 11
   While L& > 1
     Get #NX0%, L&, TD0$
     DCC& = CVS(Left$(TD0$, 4))
     If DCC& < DOCUNU& Then GoTo 9
     L& = L& - 1
   Wend
   '
9  CAPAGINAS% = 0
   For KKL& = L& To URE&
     Get #NX0%, KKL&, TD0$
     TBUF$ = TD0$
     If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
     If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
       TPSP$ = TPSP$ + Mid$(TBUF$, 5)
     End If
   Next KKL&
   '
95 SPOOLSTRING = TPSP$
   Close #NX0%
   '
10 PPXX% = InStr(TPSP$, Chr$(255))
   If PPXX% > 0 Then
      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
      TPSP$ = Mid$(TPSP$, PPXX% + 1)
      '
      If Len(COPRIAT$) > 0 Then
        COPRI% = Asc(Left$(COPRIAT$, 1))
        ATRIP$ = Mid$(COPRIAT$, 2)
        Call SHOWCOMMAND(COPRI%, ATRIP$)
      End If
      GoTo 10
      '
   End If
   '
   Screen.MousePointer = 1
   PAGINACTIVA% = 0
   ANCHPAGINA = EDIFORMU.Picture1(PAGINACTIVA%).Width
   EDIFORMU.Picture1(PAGINACTIVA%).Refresh
   EDIFORMU.VScroll1.Min = 0
   EDIFORMU.VScroll1.Value = 0
   EDIFORMU.VScroll1.Max = 1
   If ALTUPAGINA > EDIFORMU.Frame1.Height Then
      EDIFORMU.VScroll1.Max = (ALTUPAGINA - EDIFORMU.Frame1.Height) / 100
   End If
   EDIFORMU.HScroll1.Min = 0
   EDIFORMU.HScroll1.Value = 0
   EDIFORMU.HScroll1.Max = 1
   If ANCHPAGINA > EDIFORMU.Frame1.Width Then
      EDIFORMU.HScroll1.Max = (ANCHPAGINA - EDIFORMU.Frame1.Width) / 100
   End If
   '
   EDIFORMU.Show
   '
99 Command8.Enabled = True
   '
End Sub

Private Sub Command99_Click()
   Screen.MousePointer = 11
   FIN& = ULTREG&("PDOCLST")
   For U& = 73740 To FIN&
     XX$ = REGLEIDO$("PDOCLST", U&)
     Call REPLA(XX$, MKS$(U&), 1, 4)
     If U& >= 73740 Then
       If U& <= 73742 Then
         Call REPLA(XX$, "Documento Numero " + TRIM$(Str$(U&)), 5, 124)
       End If
     End If
     Call GRAREG("PDOCLST", XX$, U&)
   Next U&
   Call CIERRARCH("PDOCLST")
   Call FRESHECHO("PDOCLST")
   '
   FIN& = ULTREG&("PDOCSPL")
   TTXX$ = ""
   For U& = 1470000 To FIN&
     XX$ = REGLEIDO$("PDOCSPL", U&)
     IU& = CVS(Left$(XX$, 4))
     If IU& < 10000 Then
       IU& = IU& + 73742
       Call REPLA(XX$, MKS$(IU&), 1, 4)
       Call GRAREG("PDOCSPL", XX$, U&)
     End If
   Next U&
   Call CIERRARCH("PDOCSPL")
   Screen.MousePointer = 1
   Call COMUNI("Recuperacion Completada")
End Sub

'
Private Sub Form_Load()
    '
    Call CENTRAFORM(Me)
    Show
    Refresh
    DoEvents
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    If APLINVO$ <> "" Then ' DISEÑADOR DE FORMULARIOS
       If EXISTE%(LUDAT$ + APLINVO$ + ".SQR") > 0 Then
          If EXISTE%("DSGFRM09.EXE") > 0 Then
             Call CONECRUN("DSGFRM09/" + APLINVO$, "")
             Call FINAL("")
          End If
       End If
    End If
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
      Caption = Caption + "  -  " + EPAC$
    End If
    DoEvents
    '
    HOII% = HOY(HOS$)
    If HOII% = FECHANUM("21/11/05") Then
      Command99.Visible = True
    End If
    '
    Screen.MousePointer = 11
    Call CARLISFUEN
    Call CARINDIFORM
    '
    CODFOR$ = APLINVO$
    Text1.TEXT = CODFOR$
    Text2.TEXT = ECOARCH$("INDIFRM", CODFOR$)
    '
    If TRIM$(Text1) <> "" Then
      Call SHOWFORM
         'Else
         'Screen.MousePointer = 1
         'Call Command1_Click
    End If
    Screen.MousePointer = 1
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub HScroll1_Scroll()
   Picture1.Left = (-100) * HScroll1.Value
End Sub

Private Sub Picture1_DblClick()
  LU$ = LUDAT$
  NX% = FILEOPEN%(LU$ + CODFOR$ + ".FRM", 0, 128)
  Dim RF As String * 128
  For I% = 1 To LOF(NX%) / 128
    Get #NX%, I%, RF$
    If TRIM$(Left$(RF$, 8)) = "LINEWIN" Then
      H1 = CVS(Mid$(RF$, 57, 4))
      V1 = CVS(Mid$(RF$, 61, 4))
      H2 = CVS(Mid$(RF$, 65, 4))
      V2 = CVS(Mid$(RF$, 69, 4))
      If H1 > H2 Then H3 = H1: H1 = H2: H2 = H3
      If V1 > V2 Then V3 = V1: V1 = V2: V2 = V3
      '
      If H1 * 56.7 - 50 < XMOUSE Then
        If H2 * 56.7 + 50 > XMOUSE Then
          If V1 * 56.7 - 50 < YMOUSE Then
            If V2 * 56.7 + 50 > YMOUSE Then
              MsgBox XMOUSE & "   " & YMOUSE
            End If
          End If
        End If
      End If
    End If
  Next I%
  Close #NX%
  '
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   XMOUSE = X
   YMOUSE = Y
End Sub

Private Sub Picture2_DblClick()
   '
   Picture2.Visible = False
   Frame1.Visible = True
   VScroll1.Visible = True
   HScroll1.Visible = True
   Command1.Enabled = True
   Command2.Enabled = True
   Command3.Enabled = True
   Command4.Enabled = True
   Command5.Enabled = True
   Command6.Enabled = True
   Command7.Enabled = True
   Command8.Enabled = True
   Line2.Visible = True
   '
   Picture1.Visible = True
   Picture1.Refresh
   '
End Sub

Private Sub Text1_DblClick()
    '
    Call SELECHO("INDIFRM")
    CODFOR$ = VALACT1$("INDIFRM")
    If TRIM$(CODFOR$) = "" Then
        Text1.SetFocus
      Else
        Text1.TEXT = CODFOR$
        Text2.TEXT = VALACT2$("INDIFRM")
        Call SHOWFORM
    End If
    '
End Sub

Private Sub Text2_LOSTFOCUS()
    Call GRATITUFORM
End Sub

Sub CARLISFUEN()
    '
    NOMIMPRE$ = TRIM$(Printer.DeviceName)
    If NOMIMPRE$ = "" Then
         Call MENSERR(24, "Falta Seleccionar Impresora\en la Configuración de Windows")
         Call FINAL("")
    End If
    '
    'Call BLANARCH("FUENPRIN")
    Call SETULTREG("FUENPRIN", 0)
    On Error Resume Next
    For I% = 1 To Printer.FontCount
        Printer.FontName = Printer.Fonts(I% - 1)
        List1.AddItem Printer.FontName
    Next I%
    On Error GoTo 0
    '
    JJ& = 0
    For I% = 1 To List1.ListCount
      YY$ = TRIM$(List1.List(I% - 1))
      If YY$ <> "" Then
        If InStr(UCase$(YY$), "BENGUIAT") < 1 Then
          XX$ = REGBLAN$("FUENPRIN")
          Call REPLA(XX$, YY$, 1, 42)
          Printer.FontName = YY$
          A1 = 0: A2 = 0
          On Error Resume Next
          A1 = Printer.TextWidth(String$(80, "l"))
          A2 = Printer.TextWidth(String$(80, "M"))
          On Error GoTo 0
          If A1 > 0 And A2 > 0 Then
            If A1 = A2 Then
              Call REPLA(XX$, "(fija)", 43, 6)
            End If
            'Call REGAPP("FUENPRIN", XX$)
            JJ& = JJ& + 1
            Call GRAREG("FUENPRIN", XX$, JJ&)
          End If
        End If
      End If
    Next I%
    '
    Call SETULTREG("FUENPRIN", JJ&)
    Call CIERRARCH("FUENPRIN")
    '
End Sub

Sub CARINDIFORM()
    '
    Screen.MousePointer = 11
    Dim RF As String * 128
    File1.Path = LUDAT$
    File1.Pattern = "*.FRM"
    JJ& = 0
    '
    For U& = 0 To File1.ListCount - 1
      File1.ListIndex = U&
      ARXYZ$ = UCase$(File1.filename)
      PPXX% = InStr(ARXYZ$, ".FRM")
      If PPXX% > 1 Then
        ARXYZ$ = TRIM$(Left$(ARXYZ$, PPXX% - 1))
        If ARXYZ$ <> "" Then
          TITUFOR$ = ""
          RF11$ = LOADFILE(LUDAT$ + ARXYZ$ + ".FRM")
          If Len(RF11$) >= 128 Then
            '
            For I1& = 1 To Len(RF11$) Step 128
              RF$ = Mid$(RF11$, I1&, 128)
              If TRIM$(Left$(RF$, 8)) = "TITUFORM" Then
                TITUFOR$ = Mid$(RF$, 9, 48)
                GoTo 10
              End If
            Next I1&
            '
10          INF$ = REGBLAN$("INDIFRM")
            Call REPLA(INF$, ARXYZ$, 1, 8)
            Call REPLA(INF$, TITUFOR$, 9, 48)
            JJ& = JJ& + 1
            Call GRAREG("INDIFRM", INF$, JJ&)
            '
          End If
        End If
      End If
19  Next U&
    '
    Call SETULTREG("INDIFRM", JJ&)
    Call CIERRARCH("INDIFRM")
    Call FRESHECHO("INDIFRM")
    Screen.MousePointer = 1
    '
End Sub
'
Sub SHOWFORM()
   '
   Static CACOMM%, PRICOMM$(128)
   '
   FORMALTER.PSPOOL.Clear
   MARIZW = Val(ATRIFORM$(CODFOR$, "MARIZQMM"))
   MARSUW = Val(ATRIFORM$(CODFOR$, "MARSUPMM"))
   '
   MUESTRAFORBLA% = 1
   Call CARFORWIN
   MUESTRAFORBLA% = 0
   Call SETSPOOL("ENDDOC", "")
   '
   If CACOMM% < 1 Then
      PRICOMM$(1) = "ORIENTATION"
      PRICOMM$(2) = "NEWPAGE"
      PRICOMM$(3) = "ENDDOC"
      PRICOMM$(4) = "HEIGHT"
      PRICOMM$(5) = "WIDTH"
      PRICOMM$(6) = "SCALEMODE"
      PRICOMM$(7) = "FORECOLOR"
      PRICOMM$(8) = "DRAWMODE"
      PRICOMM$(9) = "DRAWWIDTH"
      PRICOMM$(10) = "DRAWSTYLE"
      PRICOMM$(11) = "FILLSTYLE"
      PRICOMM$(12) = "LINE"
      PRICOMM$(13) = "FONTNAME"
      PRICOMM$(14) = "FONTITALIC"
      PRICOMM$(15) = "FONTSIZE"
      PRICOMM$(16) = "FONTBOLD"
      PRICOMM$(17) = "CURRENTX"
      PRICOMM$(18) = "CURRENTY"
      PRICOMM$(19) = "PRINT"
      PRICOMM$(20) = "FONTUNDERLINE"
      PRICOMM$(21) = "PAINTPICTURE"
      CACOMM% = 21
   End If
   '
10 TPSP$ = ""
   For KKII% = 1 To FORMALTER.PSPOOL.ListCount
      TTTT$ = TRIM$(FORMALTER.PSPOOL.List(KKII% - 1))
      PCOMM$ = TRIM$(Left$(TTTT$, 16))
      PATRI$ = Mid$(TTTT$, 17)
      For KKJJ% = 1 To CACOMM%
        If PRICOMM$(KKJJ%) = PCOMM$ Then
           TPSP$ = TPSP$ + Chr$(KKJJ%) + PATRI$ + Chr$(255)
           GoTo 19
        End If
      Next KKJJ%
19 Next KKII%
   '
   Picture1.Cls
   Picture1.Top = 0
   Picture1.Refresh
   Picture1.Height = 7070
   Picture1.Width = 11700
   Picture1.DrawWidth = 3
   Picture1.DrawStyle = 0
   Picture1.FillStyle = 1
   '
   ALTUPAGINA = 7070
   TOPEPAGINA = 0
   '
   SPOOLSTRING = TPSP$
   '
20 PPXX% = InStr(TPSP$, Chr$(255))
   If PPXX% > 0 Then
      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
      TPSP$ = Mid$(TPSP$, PPXX% + 1)
      '
      If Len(COPRIAT$) > 0 Then
        COPRI% = Asc(Left$(COPRIAT$, 1))
        If COPRI% = 2 Or COPRI% = 3 Then GoTo 30
        ATRIP$ = Mid$(COPRIAT$, 2)
        Call SHOWCOMM(COPRI%, ATRIP$)
      End If
      GoTo 20
      '
   End If
   '
30 Screen.MousePointer = 1
   ANCHPAGINA = Picture1.Width
   Picture1.Refresh
   VScroll1.Min = 0
   VScroll1.Value = 0
   VScroll1.Max = 1
   If ALTUPAGINA > Frame1.Height Then
      VScroll1.Max = (ALTUPAGINA - Frame1.Height) / 100
   End If
   HScroll1.Min = 0
   HScroll1.Value = 0
   HScroll1.Max = 1
   If ANCHPAGINA > Frame1.Width Then
      HScroll1.Max = (ANCHPAGINA - Frame1.Width) / 100
   End If
   '
   Picture1.Refresh
   '
   MARIZW = Val(ATRIFORM$(CODFOR$, "MARIZQMM"))
   MARSUW = Val(ATRIFORM$(CODFOR$, "MARSUPMM"))
   DETDESMM = Val(ATRIFORM$(CODFOR$, "DETDESMM"))
   DETHASMM = Val(ATRIFORM$(CODFOR$, "DETHASMM"))
   If DETHASMM < DETDESMM + 4 + INTERLI Then
      DETHASMM = DETDESMM + 4 * INTERLI
   End If
   Picture1.ForeColor = RGB(0, 100, 0)
   Picture1.Line (56.7 * (MARIZW + 1), 56.7 * (DETDESMM + MARSUW))-(EXTREDET - 56.7, 56.7 * (DETHASMM + MARSUW)), RGB(200, 255, 255), BF
   INTERX$ = TRIM$(ATRIFORM$(CODFOR$, "INTERLIN"))
   INTERLI = Val(INTERX$)
   If INTERLI < 1 Then INTERLI = 4
   '
   AFORMX$ = LOADFILE$(LUDAT$ + CODFOR$ + ".FRM")
   LNXP& = 1 + Len(AFORMX$) / 128
   For KU& = 1 To LNXP&
     RFP$ = Mid$(AFORMX$, 128 * KU& - 127, 128)
200  FUENX$ = TRIM$(Mid$(RFP$, 87, 32))
     FORMA1$ = TRIM$(Mid$(RFP$, 75, 6))
     FORMA2$ = TRIM$(Mid$(RFP$, 81, 6))
     If TRIM$(FORMA1$ + FORMA2$) <> "" Then
       On Error Resume Next
       Picture1.FontName = FUENX$
       If Err Then
         Picture1.FontName = "Courier"
       End If
       On Error GoTo 0
       On Error Resume Next
       Picture1.FontSize = Str$(CVS(Mid$(RFP$, 119, 4)))
       On Error GoTo 0
       Picture1.FontBold = False
       If Asc(Mid$(RFP$, 123, 1)) > 32 Then
         Picture1.FontBold = True
       End If
       Picture1.FontItalic = False
       If Asc(Mid$(RFP$, 124, 1)) > 32 Then
         Picture1.FontItalic = True
       End If
       Picture1.FontUnderline = False
       If Asc(Mid$(RFP$, 125, 1)) > 32 Then
         Picture1.FontUnderline = True
       End If
       '
       PH1 = 56.7 * (CVS(Mid$(RFP$, 57, 4)) + MARIZW + X0)
       PV1 = 56.7 * (CVS(Mid$(RFP$, 61, 4)) + MARSUW + Y0)
       FORMA1$ = TRIM$(Mid$(RFP$, 75, 6))
       '
       PH2 = 56.7 * (CVS(Mid$(RFP$, 65, 4)) + MARIZW + X0)
       PV2 = 56.7 * (CVS(Mid$(RFP$, 69, 4)) + MARSUW + Y0)
       FORMA2$ = TRIM$(Mid$(RFP$, 81, 6))
       '
       If FORMA1$ <> "" Then
         TTXX$ = TRIM$(Left$(RFP$, 8))
         AMAXI% = Val(Mid$(FORMA1$, 2))
         TTXX$ = Left$(TTXX$, AMAXI%)
         If Left$(FORMA1$, 1) = "A" Or Left$(FORMA1$, 1) = "C" Or Left$(FORMA1$, 1) = "D" Then
             TTXX$ = Left$(TTXX$ + String$(256, "_"), AMAXI%)
           ElseIf Left$(FORMA1$, 1) <> "X" Then
             While Len(TTXX$) < AMAXI%
               TTXX$ = "_" + TTXX$
             Wend
         End If
         If UCase$(Left$(TRIM$(FORMA1$), 1) = "X") Then
           TTXX$ = TTXX$ + CRLF$ + TRIM$(Mid$(RFP$, 9, 40))
           PPXXZZ% = InStr(FORMA1$, "/"): If PPXXZZ% < 1 Then PPXXZZ% = 1 + Len(FORMA1$)
           CAREMA% = Val(Mid$(FORMA1$, 1 + PPXXZZ%)): If CAREMA% < 1 Then CAREMA% = 1
           INTERLIX = INTERLI: If INTERLIX / CAREMA% < 3 Then INTERLIX = INTERLI * (1 + CAREMA%)
           Call FRATEXTO(TTXX$, AMAXI%)
           For CAREI% = 1 To CARETEX%
             If CAREI% <= CAREMA% Then
               Picture1.ForeColor = RGB(255, 0, 0)
               Picture1.CurrentX = PH1
               Picture1.CurrentY = PV1
               If Picture1.Height < Picture1.CurrentY + 1000 Then
                  Picture1.Height = Picture1.CurrentY + 1000
               End If
               If CAREI% = 1 Then
                 If PV1 <= (MARSUW + 4) * 56.7 Then
                   PV1 = PV1 + 56.7 * DETDESMM
                 End If
               End If
               TTYY$ = Left$(RETEX$(CAREI%) + String$(256, "_"), AMAXI%)
               Picture1.Print TTYY$
               PV1 = PV1 + INTERLIX / (CAREMA% + 1) * 56.7
             End If
           Next CAREI%
           GoTo 880
         End If
         If OFSECOPIA% > 0 Then
           TTXY$ = TTXX$
           While Len(TTXX$) < OFSECOPIA%
             TTXX$ = TTXX$ + " "
           Wend
           TTXX$ = RTrim$(TTXX$ + TTXY$)
         End If
         If TRIM$(TTXX$) <> "" Then
           If PV1 <= (MARSUW + 4) * 56.7 Then
             PV1 = PV1 + 56.7 * DETDESMM
           End If
           Picture1.CurrentX = PH1
           Picture1.CurrentY = PV1
           If Picture1.Height < Picture1.CurrentY + 1000 Then
              Picture1.Height = Picture1.CurrentY + 1000
           End If
           If TRIM$(TTXX$) <> "" Then
              Picture1.ForeColor = RGB(255, 0, 0)
              Picture1.Print TTXX$
           End If
         End If
       End If
       '
880    If FORMA2$ <> "" Then
         TTXX$ = TRIM$(Left$(RFP$, 8))
         AMAXI% = Val(Mid$(FORMA2$, 2))
         TTXX$ = Left$(TTXX$, AMAXI%)
         If Left$(FORMA2$, 1) = "A" Or Left$(FORMA2$, 1) = "C" Or Left$(FORMA2$, 1) = "D" Then
             TTXX$ = Left$(TTXX$ + String$(256, "_"), AMAXI%)
           ElseIf Left$(FORMA2$, 1) <> "X" Then
             While Len(TTXX$) < AMAXI%
               TTXX$ = "_" + TTXX$
             Wend
         End If
         If UCase$(Left$(TRIM$(FORMA2$), 1) = "X") Then
           TTXX$ = TTXX$ + CRLF$ + TRIM$(Mid$(RFP$, 9, 40))
           PPXXZZ% = InStr(FORMA2$, "/"): If PPXXZZ% < 1 Then PPXXZZ% = 1 + Len(FORMA2$)
           CAREMA% = Val(Mid$(FORMA2$, 1 + PPXXZZ%)): If CAREMA% < 1 Then CAREMA% = 1
           INTERLIX = INTERLI: If INTERLIX / CAREMA% < 3 Then INTERLIX = INTERLI * (1 + CAREMA%)
           Call FRATEXTO(TTXX$, AMAXI%)
           For CAREI% = 1 To CARETEX%
             If CAREI% <= CAREMA% Then
               Picture1.ForeColor = RGB(255, 0, 0)
               Picture1.CurrentX = PH2
               Picture1.CurrentY = PV2
               If Picture1.Height < Picture1.CurrentY + 1000 Then
                  Picture1.Height = Picture1.CurrentY + 1000
               End If
               If CAREI% = 1 Then
                 If PV2 <= (MARSUW + 4) * 56.7 Then
                   PV2 = PV2 + 56.7 * DETDESMM
                 End If
               End If
               TTYY$ = Left$(RETEX$(CAREI%) + String$(256, "_"), AMAXI%)
               Picture1.Print TTYY$
               PV2 = PV2 + INTERLIX / (CAREMA% + 1) * 56.7
             End If
           Next CAREI%
           GoTo 890
         End If
         If OFSECOPIA% > 0 Then
           TTXY$ = TTXX$
           While Len(TTXX$) < OFSECOPIA%
             TTXX$ = TTXX$ + " "
           Wend
           TTXX$ = RTrim$(TTXX$ + TTXY$)
         End If
         If TRIM$(TTXX$) <> "" Then
           If PV2 <= (MARSUW + 4) * 56.7 Then
             PV2 = PV2 + 56.7 * DETDESMM
           End If
           Picture1.CurrentX = PH2
           Picture1.CurrentY = PV2
           If Picture1.Height < Picture1.CurrentY + 1000 Then
              Picture1.Height = Picture1.CurrentY + 1000
           End If
           If TRIM$(TTXX$) <> "" Then
              Picture1.ForeColor = RGB(255, 0, 0)
              Picture1.Print TTXX$
           End If
         End If
       End If
       '
890 End If
     '
   Next KU&
   '
   If ALTUPAGINA < Picture1.Height Then
     ALTUPAGINA = Picture1.Height
   End If
   If ALTUPAGINA > Frame1.Height Then
      VScroll1.Max = (ALTUPAGINA - Frame1.Height) / 100
   End If
End Sub

Sub SHOWCOMM(COPRI%, ATRIP$)
   '
   If PAGINACTIVA% > CAPAGINAS% Then
      CAPAGINAS% = PAGINACTIVA%
   End If
10 On Error Resume Next
   Select Case COPRI%
   Case 1
     '
   Case 2
     PAGINACTIVA% = PAGINACTIVA% + 1
   Case 3
     PAGINACTIVA% = PAGINACTIVA% + 1
   Case 4
     '
   Case 5
     Picture1.Width = Val(TRIM$(ATRIP$))
   Case 6
     Picture1.ScaleMode = Val(TRIM$(ATRIP$))
   Case 7
     Picture1.ForeColor = Val(TRIM$(ATRIP$))
   Case 8
     Picture1.DrawMode = Val(TRIM$(ATRIP$))
   Case 9
     Picture1.DrawWidth = Val(TRIM$(ATRIP$))
   Case 10
     Picture1.DrawStyle = Val(TRIM$(ATRIP$))
   Case 11
     Picture1.FillStyle = Val(TRIM$(ATRIP$))
   Case 12
     AA11 = Val(Mid$(ATRIP$, 1, 16))
     BB11 = Val(Mid$(ATRIP$, 17, 16))
     AA22 = Val(Mid$(ATRIP$, 33, 16))
     BB22 = Val(Mid$(ATRIP$, 49, 16))
     CCOLO = Val(Mid$(ATRIP$, 65, 16))
     If BB11 + 200 > ALTUPAGINA Then ALTUPAGINA = BB11 + 200
     If BB22 + 200 > ALTUPAGINA Then ALTUPAGINA = BB22 + 200
     If Picture1.Height < ALTUPAGINA - 1 Then
       Picture1.Height = ALTUPAGINA
     End If
     If AA11 > Picture1.Width - 300 Then
        Picture1.Width = AA11 + 300
     End If
     If AA22 > Picture1.Width - 300 Then
        Picture1.Width = AA22 + 300
     End If
     BBLO$ = UCase$(TRIM$(Mid$(ATRIP$, 81, 16)))
     If BBLO$ = "B" Then
        Picture1.Line (AA11, BB11 + TOPEPAGINA)-(AA22, BB22 + TOPEPAGINA), CCOLO, B
       Else
        Picture1.Line (AA11, BB11 + TOPEPAGINA)-(AA22, BB22 + TOPEPAGINA), CCOLO
     End If
   Case 13
     Picture1.FontName = TRIM$(ATRIP$)
   Case 14
     Picture1.FontItalic = Val(TRIM$(ATRIP$))
   Case 15
     Picture1.FontSize = Val(TRIM$(ATRIP$))
   Case 16
     Picture1.FontBold = Val(TRIM$(ATRIP$))
   Case 17
     If Val(TRIM$(ATRIP$)) > Picture1.Width - 300 Then
        Picture1.Width = Val(TRIM$(ATRIP$)) + 300
     End If
     Picture1.CurrentX = Val(TRIM$(ATRIP$))
   Case 18
     If Val(TRIM$(ATRIP$)) + 200 > ALTUPAGINA Then ALTUPAGINA = Val(TRIM$(ATRIP$)) + 1000
       If Picture1.Height < ALTUPAGINA - 1 Then
         Picture1.Height = ALTUPAGINA
       End If
     Picture1.CurrentY = Val(TRIM$(ATRIP$)) + TOPEPAGINA
   Case 19
     ATRIP$ = RTrim$(ATRIP$)
     If Right$(ATRIP$, 1) = ";" Then
         ATRIP$ = Left$(ATRIP$, Len(ATRIP$) - 1)
         Picture1.Print ATRIP$;
        Else
         Picture1.Print ATRIP$
     End If
   Case 20
     Picture1.FontUnderline = Val(TRIM$(ATRIP$))
   Case 21
     REDUFORMU = 1
     Call INCLUDEIMG(ATRIP$, "SHOW")
   End Select
   On Error GoTo 0
   '
End Sub

Private Sub VScroll1_Scroll()
   Picture1.Top = (-100) * VScroll1.Value
End Sub

Private Sub Command9_Click()
   '
   VScroll1.Visible = False
   HScroll1.Visible = False
   Command1.Enabled = False
   Command2.Enabled = False
   Command3.Enabled = False
   Command4.Enabled = False
   Command5.Enabled = False
   Command6.Enabled = False
   Command7.Enabled = False
   Command8.Enabled = False
   Line2.Visible = False
   '
   ALPAMAX = 0: ANPAMAX = 0
     If Picture1.Width > ANPAMAX Then
       ANPAMAX = Picture1.Width + 200
     End If
     If Picture1.Height > ALPAMAX Then
       ALPAMAX = Picture1.Height + 200
     End If
   REDUV = ScaleHeight / ALPAMAX
   REDUH = ScaleWidth / ANPAMAX
   REDUC = REDUV: If REDUH < REDUC Then REDUC = REDUH
   REDUFORMU = REDUC
   '
   Picture2.Cls
   Picture2.Top = (ScaleHeight - ALPAMAX * REDUC) / 2
   Picture2.Left = (ScaleWidth - ANPAMAX * REDUC) / 2
   Picture2.Height = ALPAMAX * REDUC
   Picture2.Width = ANPAMAX * REDUC
   Picture2.Refresh
   Picture2.DrawWidth = 2
   Picture2.DrawStyle = 0
   Picture2.FillStyle = 1
   '
   TOPEPAGINA = 0
   PAGINACTIVA% = 0
   CAPAGINAS% = 0
   '
   Screen.MousePointer = 11
   TPSP$ = SPOOLSTRING$
10 PPXX% = InStr(TPSP$, Chr$(255))
   If PPXX% > 0 Then
      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
      TPSP$ = Mid$(TPSP$, PPXX% + 1)
      '
      If Len(COPRIAT$) > 0 Then
        COPRI% = Asc(Left$(COPRIAT$, 1))
        ATRIP$ = Mid$(COPRIAT$, 2)
        Call SHOWCOMMZOOM(COPRI%, ATRIP$)
      End If
      GoTo 10
      '
   End If
   '
   Frame1.Visible = False
   Screen.MousePointer = 1
   Picture2.Refresh
      Picture1.Visible = False
      Picture2.Visible = False
   Picture2.Visible = True
   Picture2.Refresh
   '
End Sub

Sub SHOWCOMMZOOM(COPRI%, ATRIP$)
   '
   If PAGINACTIVA% > CAPAGINAS% Then
      CAPAGINAS% = PAGINACTIVA%
   End If
10 On Error Resume Next
   Select Case COPRI%
   Case 1
     'Printer.Orientation = Val(TRIM$(ATRIP$))
   Case 2
     PAGINACTIVA% = PAGINACTIVA% + 1
   Case 3
     PAGINACTIVA% = PAGINACTIVA% + 1
   Case 4
     'ALTUPAGINA = Val(TRIM$(ATRIP$))
   Case 5
     'Picture1.Width = Val(TRIM$(ATRIP$))
   Case 6
     Picture2.ScaleMode = Val(TRIM$(ATRIP$))
   Case 7
     Picture2.ForeColor = Val(TRIM$(ATRIP$))
   Case 8
     Picture2.DrawMode = Val(TRIM$(ATRIP$))
   Case 9
     Picture2.DrawWidth = 0.66 * Val(TRIM$(ATRIP$))
   Case 10
     Picture2.DrawStyle = Val(TRIM$(ATRIP$))
   Case 11
     Picture2.FillStyle = Val(TRIM$(ATRIP$))
   Case 12
     AA11 = Val(Mid$(ATRIP$, 1, 16)) * REDUFORMU
     BB11 = Val(Mid$(ATRIP$, 17, 16)) * REDUFORMU
     AA22 = Val(Mid$(ATRIP$, 33, 16)) * REDUFORMU
     BB22 = Val(Mid$(ATRIP$, 49, 16)) * REDUFORMU
     CCOLO = Val(Mid$(ATRIP$, 65, 16))
     'If BB11 + 1000 > ALTUPAGINA Then ALTUPAGINA = BB11 + 1000
     'If BB22 + 1000 > ALTUPAGINA Then ALTUPAGINA = BB22 + 1000
     'For KKI% = 0 To CAPAGINAS%
     '  If Picture1(KKI%).Height < ALTUPAGINA Then
     '    Picture1(KKI%).Height = ALTUPAGINA
     '  End If
     'Next KKI%
     'If AA11 > Picture1.Width Then
     '   Picture1.Width = AA11
     'End If
     'If AA22 > Picture1.Width Then
     '   Picture1.Width = AA22
     'End If
     BBLO$ = UCase$(TRIM$(Mid$(ATRIP$, 81, 16)))
     If BBLO$ = "B" Then
        Picture2.Line (AA11, BB11 + TOPEPAGINA)-(AA22, BB22 + TOPEPAGINA), CCOLO, B
       Else
        Picture2.Line (AA11, BB11 + TOPEPAGINA)-(AA22, BB22 + TOPEPAGINA), CCOLO
     End If
   Case 13
     TEPRUORI.FontName = TRIM$(ATRIP$)
   Case 14
     TEPRUORI.FontItalic = Val(TRIM$(ATRIP$))
   Case 15
     TEPRUORI.FontSize = Val(TRIM$(ATRIP$))
   Case 16
     TEPRUORI.FontBold = Val(TRIM$(ATRIP$))
   Case 17
     Picture2.CurrentX = Val(TRIM$(ATRIP$)) * REDUFORMU
   Case 18
     Picture2.CurrentY = (Val(TRIM$(ATRIP$)) + TOPEPAGINA) * REDUFORMU
   Case 19
     TEPRUORI.Caption = ATRIP$
     AMATEX = TEPRUORI.Width * REDUFORMU
     TEPRURED.Font = "Arial"
     If InStr(UCase$(TEPRUORI.FontName), "COURIER") > 0 Then
        TEPRURED.FontName = "Courier New"
     End If
     TEPRURED.FontBold = False
     TEPRURED.Caption = ATRIP$
     For KKKK = 24 To 1 Step -1
       TEPRURED.FontSize = KKKK / 2
       If TEPRURED.Width <= AMATEX Then GoTo 17
     Next KKKK
     '
17   Picture2.FontName = TEPRURED.FontName '"Arial"
     Picture2.FontSize = KKKK / 2
     Picture2.FontBold = False
     ATRIP$ = RTrim$(ATRIP$)
     If Right$(ATRIP$, 1) = ";" Then
         ATRIP$ = Left$(ATRIP$, Len(ATRIP$) - 1)
         Picture2.Print ATRIP$;
        Else
         Picture2.Print ATRIP$
     End If
   Case 20
     Picture2.FontUnderline = Val(TRIM$(ATRIP$))
   Case 21
     Call INCLUDEIMG(ATRIP$, "SHOW")
   End Select
   On Error GoTo 0
   '
End Sub
'
Sub INCLUDEIMG(ATRIP$, MODOPRE$)
   '
   BBUS% = Val(Mid$(ATRIP$, 1, 16))
   IMAGENUMBER& = Val(Mid$(ATRIP$, 17, 16))
   PHX1 = Val(Mid$(ATRIP$, 33, 16)) * REDUFORMU
   PVX1 = Val(Mid$(ATRIP$, 49, 16)) * REDUFORMU
   ACHX = Val(Mid$(ATRIP$, 65, 16)) * REDUFORMU
   ALTX = Val(Mid$(ATRIP$, 81, 16)) * REDUFORMU
   '
   RFF$ = RECFILT$("PIMGLST", 1, MKI$(BBUS%))
   For KKI% = 1 To Len(RFF$) Step 4
     RELI& = CVS(Mid$(RFF$, KKI%, 4))
     If RELI& > 0 Then
       ZZ$ = REGLEIDO$("PIMGLST", RELI&)
       If CVS(Mid$(ZZ$, 53, 4)) = IMAGENUMBER& Then GoTo 10
     End If
   Next KKI%
   If Len(RFF$) < 4 Then
      Call MENSERR(24, "Imagen no Encontrada - Imposible Incluir.")
      Exit Sub
   End If
   '
10 NOMBIMAG$ = TRIM$(Mid$(ZZ$, 3, 38))
   LARGIMAG& = CVS(Mid$(ZZ$, 41, 4))
   FECHIMAG# = CVD(Mid$(ZZ$, 45, 8))
   FIRSTRECORD& = CVS(Mid$(ZZ$, 57, 4))
   LASTRECORD& = CVS(Mid$(ZZ$, 61, 4))
   '
   If FIRSTRECORD& < 1 Or LASTRECORD& < FIRSTRECORD& Or LASTRECORD& > ULTREG&("PIMGSPL") Then
      Call MENSERR(24, "Imagen no Encontrada - Imposible Incluir.")
      Exit Sub
   End If
   '
   If FIRSTRECORD& > 1 Then
      XX$ = REGLEIDO$("PIMGSPL", FIRSTRECORD& - 1)
      If CVS(Left$(XX$, 4)) = IMAGENUMBER& Then
        Call MENSERR(24, "Imagen no Encontrada - Imposible Incluir.")
        Exit Sub
      End If
   End If
   '
   If LASTRECORD& < ULTREG&("PIMGSPL") Then
      XX$ = REGLEIDO$("PIMGSPL", LASTRECORD& + 1)
      If CVS(Left$(XX$, 4)) = IMAGENUMBER& Then
        Call MENSERR(24, "Imagen no Encontrada - Imposible Incluir.")
        Exit Sub
      End If
   End If
   '
   IMAGSTRI$ = String$(508 * (1 + LASTRECORD& - FIRSTRECORD&), 0)
   PINSERT& = 1
   For U& = FIRSTRECORD& To LASTRECORD&
      XX$ = REGLEIDO$("PIMGSPL", U&)
      If CVS(Left$(XX$, 4)) <> IMAGENUMBER& Then
        Call MENSERR(24, "Imagen no Encontrada - Imposible Incluir.")
        Exit Sub
      End If
      Mid$(IMAGSTRI$, PINSERT&, 508) = Mid$(XX$, 5, 508)
      PINSERT& = PINSERT& + 508
   Next U&
   '
   If Len(IMAGSTRI$) < LARGIMAG& Then
      Call MENSERR(24, "Imagen no Encontrada - Imposible Incluir.")
      Exit Sub
   End If
   IMAGSTRI$ = Left$(IMAGSTRI$, LARGIMAG&)
   '
   NX% = FreeFile
   Open LUCOM$ + NOMBIMAG$ For Binary As #NX%
   Put #NX%, , IMAGSTRI$
   Close #NX%
   '
   FORMALTER.Image3.Picture = LoadPicture("")
   FORMALTER.Image3.Picture = LoadPicture(LUCOM$ + NOMBIMAG$)
   '
   HIMAGEN = FORMALTER.Image3.Picture.Width
   VIMAGEN = FORMALTER.Image3.Picture.Height
   '
   AA11 = PHX1
   BB11 = PVX1 + TOPEPAGINA * REDUFORMU
   AA22 = AA11 + ACHX
   BB22 = BB11 + ALTX
   If BB11 + 200 > ALTUPAGINA Then ALTUPAGINA = BB11 + 200
   If BB22 + 200 > ALTUPAGINA Then ALTUPAGINA = BB22 + 200
   '
   If Picture1.Height < ALTUPAGINA - 1 Then
     Picture1.Height = ALTUPAGINA
   End If
   '
   If AA11 > Picture1.Width Then
     Picture1.Width = AA11
   End If
   '
   If AA22 > Picture1.Width Then
     Picture1.Width = AA22
   End If
   '
   If REDUFORMU = 1 Then
       Picture1.PaintPicture FORMALTER.Image3.Picture, PHX1, PVX1 + TOPEPAGINA, ACHX, ALTX
     Else
       Picture2.PaintPicture FORMALTER.Image3.Picture, PHX1, PVX1 + TOPEPAGINA * REDUFORMU, ACHX, ALTX
   End If
   '
End Sub

Sub GRATITUFORM()
   '
   Screen.MousePointer = 11
   If TRIM$(Text1.TEXT) = TRIM$(CODFOR$) Then
     If TRIM$(CODFOR$) <> "" Then
       ARXYZ$ = TRIM$(CODFOR$)
       TITUFOR$ = Text2.TEXT
       NX% = FILEOPEN%(LUDAT$ + ARXYZ$ + ".FRM", 0, 128)
       Dim RF As String * 128
       '
       For I% = 1 To LOF(NX%) / 128
         Get #NX%, I%, RF$
         If TRIM$(Left$(RF$, 8)) = "TITUFORM" Then
           RFY$ = RF$
           Call REPLA(RFY$, TITUFOR$, 9, 48)
           LSet RF$ = RFY$
           Put #NX%, I%, RF$
           GoTo 10
         End If
       Next I%
       '
       I% = 1 + LOF(NX%) / 128
       RFY$ = Space$(56) + String$(72, 0)
       Call REPLA(RFY$, "TITUFORM", 1, 8)
       Call REPLA(RFY$, TITUFOR$, 9, 48)
       LSet RF$ = RFY$
       Put #NX%, I%, RF$
       '
10     Close #NX%: NX% = 0
       '
       For V& = 1 To ULTREG&("INDIFRM")
         INF$ = REGLEIDO$("INDIFRM", V&)
         If Left$(INF$, 8) = AJUSTI$(ARXYZ$, 8) Then
           If TITUFOR$ <> "" Then
             Call REPLA(INF$, TITUFOR$, 9, 48)
             Call GRAREG("INDIFRM", INF$, V&)
           End If
           GoTo 19
         End If
       Next V&
       '
       INF$ = REGBLAN$("INDIFRM")
       Call REPLA(INF$, ARXYZ$, 1, 8)
       Call REPLA(INF$, TITUFOR$, 9, 48)
       Call REGAPP("INDIFRM", INF$)
       '
     End If
   End If
   '
19 Call CIERRARCH("INDIFRM")
   Call FRESHECHO("INDIFRM")
   Screen.MousePointer = 1
   '
End Sub
'
Sub MUESTRA_DOCU(DOCUNU&)
    '
5   For KKI% = 0 To 5
      EDIFORMU.Picture1(KKI%).Cls
      EDIFORMU.Picture1(KKI%).DrawWidth = 3
      EDIFORMU.Picture1(KKI%).DrawStyle = 0
      EDIFORMU.Picture1(KKI%).FillStyle = 1
    Next KKI%
    '
    Dim TD0 As String * 512
    NX0% = FILEOPEN%(LUDAT$ + "PDOCSPL.DAT", 0, 512)
    Get #NX0%, 1, TD0$
    If CVS(Left$(TD0$, 4)) > DOCUNU& Then
      Close #NX0%
      NX0% = FILEOPEN%(LUDAT$ + "PDOCSPL.DT0", 0, 512)
    End If
    '
    URE& = Abs(LOF(NX0%) / 512)
    LI& = 0: LS& = URE&
6   E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
    L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
    Get #NX0%, L&, TD0$
    DCC& = CVS(Left$(TD0$, 4))
    If DCC& < DOCUNU& Then LI& = L&: GoTo 6
    If DCC& > DOCUNU& Then LS& = L&: GoTo 6
    GoTo 8
    '
7   Screen.MousePointer = 1
    Call MENSERR(24, "Documento no Registrado")
    GoTo 99
    '
8   TPSP$ = ""
    Screen.MousePointer = 11
    While L& > 1
      Get #NX0%, L&, TD0$
      DCC& = CVS(Left$(TD0$, 4))
      If DCC& < DOCUNU& Then GoTo 9
      L& = L& - 1
    Wend
    '
9   CAPAGINAS% = 0
    For KKL& = L& To URE&
      Get #NX0%, KKL&, TD0$
      TBUF$ = TD0$
      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
      End If
    Next KKL&
    '
95  Close #NX0%
    SPOOLSTRING$ = TPSP$
    For KKI% = 0 To 5
      EDIFORMU.Picture1(KKI%).Cls
      EDIFORMU.Picture1(KKI%).Top = 0
      EDIFORMU.Picture1(KKI%).Refresh
      EDIFORMU.Picture1(KKI%).Height = 7170
      EDIFORMU.Picture1(KKI%).Width = 11700
      EDIFORMU.Picture1(KKI%).DrawWidth = 3
      EDIFORMU.Picture1(KKI%).DrawStyle = 0
      EDIFORMU.Picture1(KKI%).FillStyle = 1
    Next KKI%
    ALTUPAGINA = 7170
    TOPEPAGINA = 0
    PAGINACTIVA% = 0
    CAPAGINAS% = 0
    '
10  PPXX% = InStr(TPSP$, Chr$(255))
    If PPXX% > 0 Then
      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
      TPSP$ = Mid$(TPSP$, PPXX% + 1)
      Call TRACE(20, FIN& - Len(TPSP$), FIN&)
      '
      If Len(COPRIAT$) > 0 Then
        COPRI% = Asc(Left$(COPRIAT$, 1))
        ATRIP$ = Mid$(COPRIAT$, 2)
        Call SHOWCOMMAND(COPRI%, ATRIP$)
      End If
      GoTo 10
      '
    End If
    '
    Screen.MousePointer = 1
    PAGINACTIVA% = 0
    For KKI% = 0 To 5
       EDIFORMU.Picture1(KKI%).Visible = False
    Next KKI%
    EDIFORMU.Picture1(PAGINACTIVA%).Visible = True
    EDIFORMU.Picture1(PAGINACTIVA%).Refresh
    EDIFORMU.VScroll1.Min = 0
    EDIFORMU.VScroll1.Value = 0
    EDIFORMU.VScroll1.Max = 1
    If ALTUPAGINA > EDIFORMU.Frame1.Height Then
       EDIFORMU.VScroll1.Max = (ALTUPAGINA - EDIFORMU.Frame1.Height) / 100
    End If
    EDIFORMU.Command3.Enabled = False
    EDIFORMU.Command5.Enabled = False
    '
    On Error Resume Next
    EDIFORMU.Show 1
    On Error GoTo 0

99  End Sub



