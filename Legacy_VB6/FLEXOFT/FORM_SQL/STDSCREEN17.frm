VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form STDSCREEN17 
   BackColor       =   &H00DAE4EB&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MUESTRA"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7830
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   7830
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2520
      Top             =   1470
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2280
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   4515
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   22
      Top             =   1470
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture22 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   2115
      Left            =   3885
      ScaleHeight     =   2115
      ScaleWidth      =   2955
      TabIndex        =   21
      ToolTipText     =   "Click con Boton Derecho para Cambiar la Imagen"
      Top             =   2205
      Visible         =   0   'False
      Width           =   2955
   End
   Begin VB.PictureBox Picture8 
      BackColor       =   &H009BBBC6&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   7410
      TabIndex        =   4
      Top             =   0
      Width           =   7410
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
         Height          =   480
         Left            =   6220
         Picture         =   "STDSCREEN17.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Envia Documento por Mail"
         Top             =   0
         Width           =   600
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1240
         TabIndex        =   29
         Text            =   "Combo1"
         Top             =   105
         Width           =   2175
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
         Height          =   480
         Left            =   3465
         Picture         =   "STDSCREEN17.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Imprime Muestra del Formulario Activo"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   600
      End
      Begin VB.PictureBox Picture3 
         Height          =   615
         Left            =   4935
         ScaleHeight     =   555
         ScaleWidth      =   690
         TabIndex        =   23
         Top             =   0
         Width           =   750
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Y:"
            Height          =   255
            Left            =   -105
            TabIndex        =   24
            Top             =   210
            Width           =   315
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "X:"
            Height          =   255
            Left            =   -105
            TabIndex        =   26
            Top             =   0
            Width           =   315
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   165
            Left            =   105
            TabIndex        =   25
            Top             =   210
            Width           =   525
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "0.0"
            Height          =   255
            Left            =   105
            TabIndex        =   27
            Top             =   0
            Width           =   525
         End
      End
      Begin VB.CommandButton PAGANT1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   0
         Picture         =   "STDSCREEN17.frx":0974
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Guarda Cambios en Formulario Activo y Presenta Pagina Anterior"
         Top             =   0
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.CommandButton PAGSIG1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   590
         Picture         =   "STDSCREEN17.frx":0DB6
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Guarda Cambios en Formulario Activo  y Presenta Pagina Siguiente"
         Top             =   0
         Visible         =   0   'False
         Width           =   600
      End
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
         Height          =   480
         Left            =   5625
         Picture         =   "STDSCREEN17.frx":11F8
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Guarda Cambios en Formulario Activo"
         Top             =   0
         Width           =   600
      End
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
         Height          =   480
         Left            =   6825
         Picture         =   "STDSCREEN17.frx":1502
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Termina - Cierra la Aplicación"
         Top             =   0
         Width           =   600
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   3765
         Picture         =   "STDSCREEN17.frx":164C
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   45
         Width           =   1515
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   3465
         X2              =   3465
         Y1              =   0
         Y2              =   480
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   3780
         X2              =   7350
         Y1              =   465
         Y2              =   465
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   3465
         X2              =   7350
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   1155
         X2              =   1155
         Y1              =   0
         Y2              =   480
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   3615
      Left            =   0
      ScaleHeight     =   3555
      ScaleWidth      =   2115
      TabIndex        =   0
      Top             =   480
      Width           =   2175
      Begin VB.HScrollBar HScroll1 
         Height          =   225
         LargeChange     =   100
         Left            =   120
         TabIndex        =   2
         Top             =   3120
         Width           =   1380
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   960
         LargeChange     =   100
         Left            =   1680
         TabIndex        =   1
         Top             =   480
         Width           =   225
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3120
         Left            =   0
         ScaleHeight     =   3120
         ScaleWidth      =   1620
         TabIndex        =   3
         Top             =   0
         Width           =   1620
         Begin VB.CommandButton Command99 
            BackColor       =   &H000000FF&
            Caption         =   "."
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   3.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1500
            Index           =   0
            Left            =   0
            MaskColor       =   &H000000FF&
            TabIndex        =   20
            Top             =   0
            Visible         =   0   'False
            Width           =   180
         End
         Begin VB.CommandButton PAGSIG 
            Caption         =   "Siguiente"
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
            Left            =   840
            Picture         =   "STDSCREEN17.frx":356E
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Guarda Cambios en Formulario Activo  y Presenta Pagina Siguiente"
            Top             =   2280
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.CommandButton PAGANT 
            Caption         =   "Anterior"
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
            Left            =   -600
            Picture         =   "STDSCREEN17.frx":39B0
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Guarda Cambios en Formulario Activo y Presenta Pagina Anterior"
            Top             =   2280
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.TextBox Text99M 
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   3.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   150
            Index           =   0
            Left            =   0
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   11
            Text            =   "STDSCREEN17.frx":3DF2
            Top             =   0
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox Text99 
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   3.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   150
            Index           =   0
            Left            =   120
            TabIndex        =   9
            Text            =   "Text1"
            Top             =   360
            Visible         =   0   'False
            Width           =   975
         End
         Begin MSFlexGridLib.MSFlexGrid GRID99 
            Height          =   1365
            Index           =   0
            Left            =   0
            TabIndex        =   12
            Top             =   0
            Visible         =   0   'False
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   2408
            _Version        =   393216
            Rows            =   28
            Cols            =   4
            FixedCols       =   0
            BackColorFixed  =   8433840
            BackColorSel    =   -2147483643
            ForeColorSel    =   12582912
            BackColorBkg    =   16777215
            FocusRect       =   2
            AllowUserResizing=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Line Line2 
            BorderStyle     =   4  'Dash-Dot
            BorderWidth     =   2
            DrawMode        =   1  'Blackness
            Visible         =   0   'False
            X1              =   120
            X2              =   1440
            Y1              =   1920
            Y2              =   1320
         End
         Begin VB.Label ECO99 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   14
            Top             =   0
            Visible         =   0   'False
            Width           =   45
         End
         Begin VB.Label Label99 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label2"
            Height          =   165
            Index           =   0
            Left            =   120
            TabIndex        =   8
            Top             =   120
            Visible         =   0   'False
            Width           =   495
         End
      End
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Label6"
      Height          =   195
      Left            =   2520
      TabIndex        =   30
      Top             =   3885
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label PAGIN 
      Caption         =   "PAGIN"
      Height          =   375
      Left            =   6120
      TabIndex        =   19
      Top             =   1440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "Label17"
      Height          =   195
      Left            =   3255
      TabIndex        =   13
      Top             =   2160
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label APRO 
      Caption         =   "APRO"
      Height          =   375
      Left            =   6120
      TabIndex        =   10
      Top             =   1080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   945
      Visible         =   0   'False
      Width           =   1695
   End
End
Attribute VB_Name = "STDSCREEN17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public PINHOR, PINVER
'
Dim YAENTRO%
Dim CANCELA%
Dim SENTIOR%(256)
Dim ARBUS$, ABUSQUE$
Dim UGRID%, UTEMU%
Dim BANDENTRAPASS%
'
Dim RS As ADODB.Recordset
'
Dim XMOUSE, YMOUSE

Dim CAGRIDS%, OTRAHOJA%, PRIREGRID%(128), CATEMUL%, PRIRETEM%(128), PX, PY, ALINGRI, COEFIHOR
Dim RECARGA%

Private Sub Command12_Click()

End Sub

Private Sub Command2_Click()
   Label6 = "" 'nombre de adjunto para mail
   If UCase$(Left$(App.EXEName, 6)) = "DSGFRM" Then
      Unload Me
      Exit Sub
   End If
   
   Hide
End Sub

Sub Command4_Click()
   '
   ORIANT = Printer.Orientation
   '
   PRINTERPORT$ = Combo1.TEXT
   On Error Resume Next
   PDVN$ = Printer.DeviceName
   On Error GoTo 0
   '
   If Combo1.TEXT <> PDVN$ Then
     Printer.TrackDefault = False
     Call SETPRINPORT(Combo1.TEXT)
     PRINTERPORT$ = Combo1.TEXT
   End If
   '
   CAGRIDS% = GRID99.Count
   For KKI% = 1 To CAGRIDS%
      PRIREGRID(KKI%) = 1
   Next KKI%
   '
   CATEMUL% = Text99M.Count
   For KKI% = 1 To CATEMUL%
      PRIRETEM(KKI%) = 1
   Next KKI%
   '
18 On Error Resume Next
   If Left$(UCase$(FOR_ORIE$), 3) = "HOR" Then
        Printer.Orientation = 2
      Else
        Printer.Orientation = 1
   End If
   On Error GoTo 0
   '
19 UGRID% = 0: UTEMU% = 0: OTRAHOJA% = 0
   VUELTA% = 1                                     ' NO IMPRIME LINEAS Y CAJAS
   '
20 For i% = 1 To CANOBJ%
      Call PRINOBJ(i%, VUELTA%)
   Next i%
   If VUELTA% = 1 Then
      VUELTA% = 2
      GoTo 20
   End If
   '
   If OTRAHOJA% > 0 Then
      Printer.NewPage
      GoTo 19
   End If
   '
   If PRINTERCONTROL% = 9999 Then
        Printer.NewPage
      Else
        Printer.EndDoc
        On Error Resume Next
        Printer.Orientation = ORIANT
        On Error GoTo 0
   End If
   '
   If Combo1.TEXT <> PDVN$ Then
     Call SETPRINPORT("RESTORE")
     PRINTERPORT$ = ""
   End If
End Sub

Sub PRINOBJ(i%, VUELTA%)
   '
   MODOX$ = TIOBJ$(i%)
   TIVUEL% = 1: If MODOX$ = "LINE" Or MODOX$ = "BOX" Then TIVUEL% = 2
   If TIVUEL% <> VUELTA% Then Exit Sub
   On Error Resume Next
   Select Case MODOX$
      Case "LABEL"
'GoTo 99
         Printer.CurrentY = 56.7 * (V1(i%) + FOR_MGSU)
         Printer.CurrentX = 56.7 * (H1(i%) + FOR_MGIZ)
         Printer.Font = FUENTE$(i%)
         Printer.FontSize = TAFUEN(i%)
         Printer.FontBold = NEGRI(i%)
         Printer.FontItalic = ITALI(i%)
         Printer.FontUnderline = SUBRA(i%)
         If INVIS(i%) <> 1 Then
            If MULIN(i%) <> 1 Then
                 Printer.Print TEXTO(i%)
              Else
                 GoTo 17         ' LO TRATA COMO UN PARAMETRO
            End If
         End If
         '
      Case "PARAM"
17       Printer.CurrentY = 56.7 * (V1(i%) + FOR_MGSU)
         Printer.CurrentX = 56.7 * (H1(i%) + FOR_MGIZ)
         Printer.Font = FUENTE$(i%)
         Printer.FontSize = TAFUEN(i%)
         Printer.FontBold = NEGRI(i%)
         Printer.FontItalic = ITALI(i%)
         Printer.FontUnderline = SUBRA(i%)
         If INVIS(i%) <> 1 Then
            If MULIN(i%) <> 1 Then
                 Printer.Print FORMATNUM$(TEXTO(i%), FORPA$(i%))
               Else
                 On Error GoTo 0
                 Call FRATEXFONT(CStr(TEXTO(i%)), H2(i%) - H1(i%), FUENTE$(i%), TAFUEN(i%), NEGRI(i%))
                 For KKI% = 1 To CARETEX%
                    If Printer.CurrentY + FORMUPLA.AutoAnch.Height < 56.7 * (V2(i%) + FOR_MGSU) Then
                       Printer.CurrentX = 56.7 * (H1(i%) + FOR_MGIZ)
                          Printer.FontSize = 1
                       Printer.Font = FUENTE$(i%)
                       Printer.FontSize = TAFUEN(i%)
                       Printer.FontBold = NEGRI(i%)
                       Printer.FontItalic = ITALI(i%)
                       Printer.FontUnderline = SUBRA(i%)
                       Printer.Print RETEX$(KKI%);
                       Printer.CurrentY = Printer.CurrentY + FORMUPLA.AutoAnch.Height
                    End If
                 Next KKI%
            End If
         End If
      Case "TEXT"
'GoTo 99
         Printer.CurrentY = 56.7 * (V1(i%) + FOR_MGSU)
         Printer.CurrentX = 56.7 * (H1(i%) + FOR_MGIZ)
         Printer.Font = FUENTE$(i%)
         Printer.FontSize = TAFUEN(i%)
         Printer.FontBold = NEGRI(i%)
         Printer.FontItalic = ITALI(i%)
         Printer.FontUnderline = SUBRA(i%)
         '
         If INVIS(i%) <> 1 Then    ' Printer.Print FORMATNUM$(TEXTO(I%), FORPA$(I%))
            If MULIN(i%) <> 1 Then
                 Printer.Print FORMATNUM$(TEXTO(i%), FORPA$(i%))
               Else
                 UTEMU% = UTEMU% + 1
                 TTXX$ = CStr(TEXTO(i%))
                 If Left$(FORPA$(i%), 1) = "X" Then
                    Moneda$ = DEMONECO$(MONEMI%)
                    TTXX$ = TEXTFORM$(TTXX$, FORPA$(i%), Moneda$)
                 End If
                 If Left$(FORPA$(i%), 1) = "Z" Then
'                    MONEDA$ = DEMONECO$(MONEMI%)
                    TTXX$ = TEXTFORM$(TTXX$, FORPA$(i%), "", "", -1) 'FORMATO Z NO IMPRIME LA MONEDA EN EL IMPORTE EN LETRAS
                 End If
                 Call FRATEXFONT(TTXX$, H2(i%) - H1(i%), FUENTE$(i%), TAFUEN(i%), NEGRI(i%))
                    If SEVOLF$ = "D443-E112" Then
                       If Printer.FontSize < 4 Then
                          MsgBox "Tamaño de Fuente " & Printer.FontSize & " Ajustado a 8"
                          Printer.FontSize = 8
                       End If
                    End If
                 For KKI% = PRIRETEM%(UTEMU%) To CARETEX%
                    If Printer.CurrentY + FORMUPLA.AutoAnch.Height < 56.7 * (V2(i%) + FOR_MGSU) Then
                         Printer.CurrentX = 56.7 * (H1(i%) + FOR_MGIZ)
                         Printer.Font = FUENTE$(i%)
                            Printer.FontSize = 1
                         Printer.FontSize = TAFUEN(i%)
                         Printer.FontBold = NEGRI(i%)
                         Printer.FontItalic = ITALI(i%)
                         Printer.FontUnderline = SUBRA(i%)
                         Printer.Print RETEX$(KKI%);
                         Printer.CurrentY = Printer.CurrentY + FORMUPLA.AutoAnch.Height
                       Else
                        PRIRETEM%(UTEMU%) = KKI%
                        OTRAHOJA% = 1
                        GoTo 99
                    End If
                 Next KKI%
            End If
         End If
         '
         iobx% = IOBJETX%(i%)
         If iobx% >= 0 Then
            If ECO99(iobx%) <> "" Then
               Printer.CurrentY = 56.7 * (V1(i%) + FOR_MGSU)
               Printer.CurrentX = ECO99(iobx%).Left + 56.7 * FOR_MGIZ
               Printer.Print ECO99(iobx%).Caption
            End If
         End If
         '
      Case "LINE"
         Printer.ScaleMode = 1
         Printer.ForeColor = 0
         Printer.DrawMode = 13
         Printer.DrawWidth = 3 * XVALO(TAFUEN$(i%))
         Printer.DrawStyle = 0
         If INVIS(i%) <> 1 Then Printer.Line (56.7 * (H1(i%) + FOR_MGIZ), 56.7 * (V1(i%) + FOR_MGSU))-(56.7 * (H2(i%) + FOR_MGIZ), 56.7 * (V2(i%) + FOR_MGSU)), QBColor(0)
      Case "BOX"
         Printer.ScaleMode = 1
         Printer.ForeColor = 0
         Printer.DrawMode = 13
         Printer.DrawWidth = 3 * XVALO(TAFUEN$(i%))
         Printer.DrawStyle = 0
         Printer.FillStyle = 1
         If INVIS(i%) <> 1 Then Printer.Line (56.7 * (H1(i%) + FOR_MGIZ), 56.7 * (V1(i%) + FOR_MGSU))-(56.7 * (H2(i%) + FOR_MGIZ), 56.7 * (V2(i%) + FOR_MGSU)), QBColor(0), B
      Case "GRAPH"
         If INVIS(i%) <> 1 Then
            If TEXTO(i%) <> "" Then
              ARIMAG$ = LUDAT$ + TEXTO(i%)
              If EXISTE%(ARIMAG$) < 1 Then ARIMAG$ = App.Path + "\" + TEXTO(i%)
              If EXISTE%(ARIMAG$) < 1 Then ARIMAG$ = TEXTO(i%)
              If EXISTE%(ARIMAG$) > 0 Then
                 Call DIBUIMAG(Printer, ARIMAG$, H1(i%) + FOR_MGIZ, H2(i%) + FOR_MGIZ, V1(i%) + FOR_MGSU, V2(i%) + FOR_MGSU)
              End If
            End If
         End If
      Case "GRID"
         If INVIS(i%) <> 1 Then
            UGRID% = UGRID% + 1
            If UGRID% > GRID99.Count - 1 Then GoTo 99
            '
            Printer.ScaleMode = 1
            Printer.ForeColor = 0
            Printer.DrawMode = 13
            Printer.DrawWidth = 1 ' * XVALO(TAFUEN$(I%))
            Printer.DrawStyle = 0
            Printer.FillStyle = 1
            '
            ' CALCULA COEFICIENTE ESTIRAMIENTO HORIZONTAL
            SUMACOL = 0
' POSIBLE ERROR 2
            For COLI& = 1 To GRID99(UGRID%).Cols
               SUMACOL = SUMACOL + GRID99(UGRID%).ColWidth(COLI& - 1)
            Next COLI&
            COEFIHOR = 1
            If SUMACOL > 0 Then COEFIHOR = (56.7 * (H2(i%) - H1(i%))) / SUMACOL
            '
            'IMPRIME TEXTOS
            PY0 = 35: PY = PY0
' POSIBLE ERROR 3
            PRIGRIDACT% = PRIREGRID%(UGRID%)
            BOSUP = 0
            For ROWI& = PRIREGRID%(UGRID%) To CANREGRID%(UGRID%)   ' GRID99(UGRID%).Rows
               If PY + GRID99(UGRID%).RowHeight(ROWI& - 1) > PY0 + 56.7 * (V2(i%) - V1(i%)) Then
                  PRIREGRID%(UGRID%) = ROWI&
                  OTRAHOJA% = 1
                  GoTo 50
               End If
               If ROWI& = PRIREGRID%(UGRID%) And ROWI& > 1 Then
                  Call PRILIGRID(i%, UGRID%, 1, BOSUP)
                  BOSUP = BOSUP + GRID99(UGRID%).RowHeight(0)
               End If
               Call PRILIGRID(i%, UGRID%, ROWI&, BOSUP)
               BOSUP = BOSUP + GRID99(UGRID%).RowHeight(ROWI& - 1)
            Next ROWI&
            '
50          ' DIBUJA LA CAJA
            Printer.Line (56.7 * (H1(i%) + FOR_MGIZ), 56.7 * (V1(i%) + FOR_MGSU))-(56.7 * (H2(i%) + FOR_MGIZ), 56.7 * (V2(i%) + FOR_MGSU)), QBColor(0), B
            '
            ' DIBUJA LINEAS VERTICALES
            PX = 0
' POSIBLE ERROR 4
            For COLI& = 1 To GRID99(UGRID%).Cols - 1
               PX = PX + COEFIHOR * GRID99(UGRID%).ColWidth(COLI& - 1)
               If PX <= 56.7 * (H2(i%) - H1(i%)) Then
                  Printer.Line (56.7 * (H1(i%) + FOR_MGIZ) + PX, 56.7 * (V1(i%) + FOR_MGSU))-(56.7 * (H1(i%) + FOR_MGIZ) + PX, 56.7 * (V2(i%) + FOR_MGSU)), QBColor(0)
               End If
            Next COLI&
            '
            ' DIBUJA LINEAS HORIZONTALES
            PY = GRID99(UGRID%).RowHeight(0)
            Printer.Line (56.7 * (H1(i%) + FOR_MGIZ), 56.7 * (V1(i%) + FOR_MGSU) + PY)-(56.7 * (H2(i%) + FOR_MGIZ), 56.7 * (V1(i%) + FOR_MGSU) + PY), QBColor(0)
            For ROWI& = PRIGRIDACT% To GRID99(UGRID%).Rows
              If ROWI& > 1 Then
                 PY = PY + GRID99(UGRID%).RowHeight(ROWI& - 1)
                 If PY <= 56.7 * (V2(i%) - V1(i%)) Then
                    Printer.Line (56.7 * (H1(i%) + FOR_MGIZ), 56.7 * (V1(i%) + FOR_MGSU) + PY)-(56.7 * (H2(i%) + FOR_MGIZ), 56.7 * (V1(i%) + FOR_MGSU) + PY), QBColor(0)
                 End If
              End If
            Next ROWI&
            '
         End If
   End Select
   Call FRESHALL
99 On Error GoTo 0
   '
End Sub

Sub PRILIGRID(i%, UGRID%, ROWI&, BOSUP)
     '
     PX = 30
     ' POSIBLE ERROR 5
     GRID99(UGRID%).Row = ROWI& - 1
     '
     For COLI& = 1 To GRID99(UGRID%).Cols
        'If PX + GRID99(UGRID%).ColWidth(COLI& - 1) <= 56.7 * (H2(I%) - H1(I%)) Then
            GRID99(UGRID%).COL = COLI& - 1
            ALIG% = GRID99(UGRID%).ColAlignment(COLI& - 1)
            '
            GRID99(UGRID%).CellFontName = FUENTE$(i%)
               GRID99(UGRID%).CellFontSize = 1
            GRID99(UGRID%).CellFontSize = TAFUEN(i%)
            GRID99(UGRID%).CellFontBold = NEGRI(i%)
            GRID99(UGRID%).CellFontItalic = ITALI(i%)
            GRID99(UGRID%).CellFontUnderline = SUBRA(i%)
            '
            VVXX$ = GRID99(UGRID%).TextMatrix(ROWI& - 1, COLI& - 1)
            Label17.Font = GRID99(UGRID%).CellFontName
            Label17.FontSize = GRID99(UGRID%).CellFontSize
            Label17.FontBold = GRID99(UGRID%).CellFontBold
               If ROWI& = 1 Then Label17.FontBold = True
            Label17.FontItalic = GRID99(UGRID%).CellFontItalic
            Label17.FontUnderline = GRID99(UGRID%).CellFontUnderline
            Label17.Caption = VVXX$
            If GRID99(UGRID%).ColWidth(COLI& - 1) < 0.05 Then VVXX$ = ""
            '
            If ALIG% = 7 Then   ' JUSTIFICADO A DERECHA
                VVXX$ = TRIM$(VVXX$)
                Label17.Caption = VVXX$ + " "
                While Label17.Width < GRID99(UGRID%).ColWidth(COLI& - 1) - 30
                   VVXX$ = " " + VVXX$
                   Label17.Caption = VVXX$ + " "
                Wend
                DESPLA = -50
              Else
                While VVXX$ <> "" And Label17.Width >= GRID99(UGRID%).ColWidth(COLI& - 1) - 30
                  VVXX$ = Left$(VVXX$, Len(VVXX$) - 1)
                  Label17.Caption = VVXX$
                Wend
                DESPLA = 50
            End If
            '
            Label17 = RTrim(Label17)     ' PARA QUITAL EL BLANCO A LA DERECHA
            If PX + Label17.Width > 56.7 * (H2(i%) - H1(i%)) Then GoTo 29
            '
            ARENGRI = GRID99(UGRID%).RowHeight(ROWI& - 1)
            ATEXTO = Label17.Height
            PZ = 0:    If ARENGRI > ATEXTO + 2 * PY Then PZ = (ARENGRI - ATEXTO - 2 * PY) / 2
            '
            Printer.CurrentY = 56.7 * (V1(i%) + FOR_MGSU) + PY + PZ
            Printer.CurrentX = 56.7 * (H1(i%) + FOR_MGIZ) + PX + DESPLA
            Printer.Font = FUENTE$(i%)        ' GRID99(UGRID%).CellFontName
               Printer.FontSize = 1
            Printer.FontSize = TAFUEN(i%)     ' GRID99(UGRID%).CellFontSize
            Printer.FontBold = NEGRI(i%)      ' GRID99(UGRID%).CellFontBold
            Printer.FontItalic = ITALI(i%)    ' GRID99(UGRID%).CellFontItalic
            Printer.FontUnderline = SUBRA(i%) ' GRID99(UGRID%).CellFontUnderline
            If ROWI& = 1 Then Printer.FontBold = True
            '
            'If Left$(FORCAMPO$(COLI&), 3) = "JPG" Then
            With GRID99(UGRID%)
               .Row = ROWI& - 1
               .COL = COLI& - 1
               GRCP = .CellPicture
               If GRCP <> 0 Then
                  ENCX$ = LTrim(.TextMatrix(.Row, .COL))
                  If UCase$(Left$(ENCX$, 8)) = "///TXT =" Then
                     Call PRIMULIN(i%, GRID99(UGRID%), .Row, .COL, BOSUP)
                     ' Call SHOWREP07.IMPRIMEMULIN(GRID99(UGRID%), .Row, .Col)
                     FINSUBTO% = ATOT%
                     GoTo 29
                  End If
                  ' ACOMODA LA POSICIÓN DE LA IMPRESORA
                  PX1 = 56.7 * (H1(i%) + FOR_MGIZ) + COEFIHOR * (.ColPos(.COL) - .ColPos(0)) ' NO SE POR QUE PERO ASI FUNCIONA ...  .Col)
                  PX2 = PX1 + COEFIHOR * .ColWidth(.COL)
                  '
                  PY1 = 56.7 * (V1(i%) + FOR_MGSU) + BOSUP   ' .RowPos(.Row)
                  PY2 = PY1 + .RowHeight(.Row)
                  '
                  Printer.PaintPicture .CellPicture, PX1, PY1, PX2 - PX1, PY2 - PY1
                  Printer.CurrentY = PY1    ' PARA RESTAURAR LA POSICION VERTICAL
                  FINSUBTO% = ATOT%
                Else
                  Printer.Print VVXX$;      ' AGREGADO EPB 17/01/2011 PARA IMPRIMIR POR CAMPO
               End If
29             PX = PX + COEFIHOR * GRID99(UGRID%).ColWidth(COLI& - 1)
            End With
     Next COLI&
     PY = PY + GRID99(UGRID%).RowHeight(ROWI& - 1)
     '
End Sub

Sub PRIMULIN(i%, NGRILLA, RWX&, CLX&, BOSUP)
     CURRX = Printer.CurrentX
     CURRY = Printer.CurrentY
     With NGRILLA
        TTXX$ = Mid$(LTrim(.TextMatrix(RWX&, CLX&)), 9)
        PX1 = 56.7 * (H1(i%) + FOR_MGIZ) + COEFIHOR * (.ColPos(CLX&) - .ColPos(0)) ' NO SE POR QUE PERO ASI FUNCIONA ...  .Col)
        PY1 = 56.7 * (V1(i%) + FOR_MGSU) + BOSUP
        Printer.CurrentY = PY1
        '
        LOTEMM = (.ColWidth(CLX&) - 30) / 56.7
        '
        FONAM = FUENTE$(i%)           ' Printer.FontName
        FOSIZ = TAFUEN(i%)            ' Printer.FontSize
        FOBOL = NEGRI(i%)             ' Printer.FontBold
        Call FRATEXFONT(TTXX$, LOTEMM, FONAM, FOSIZ, FOBOL)
        '
        For KKII% = 1 To CARETEX%
           '
           Printer.FontName = FUENTE$(i%)
              Printer.FontSize = 1
           Printer.FontSize = TAFUEN(i%)
           Printer.FontBold = NEGRI(i%)
           Printer.FontItalic = ITALI(i%)
           Printer.FontUnderline = SUBRA(i%)
           '
           Printer.CurrentX = PX1    ' CURRX
           Printer.Print RETEX$(KKII%)
           If Printer.CurrentY - CURRY > ALTUREMAX Then
              ALTUREMAX = Printer.CurrentY - CURRY
           End If
        Next KKII%
        Printer.CurrentY = CURRY
        Printer.CurrentX = CURRX + COEFIHOR * .ColWidth(CLX&)   ' Printer.TextWidth(Space$(ACOLUM%(COLU& + 1)))
     End With
     '
End Sub

Private Sub Command5_Click()
   '
   COLOMARCA = RGB(255, 100, 100)
   PRICAMOBLI% = 0
   '
   Screen.MousePointer = 11
   MODEMO% = 0
   On Error Resume Next
   MODEMO% = SECUPROC09.Check1.Value
   On Error GoTo 0
   '
   For i% = 1 To CANOBJ%
     If TIOBJ$(i%) = "TEXT" Then
       If OBLIGA%(i%) > 0 Then
         TTXX$ = TEXTO(i%)
         If TRIM$(TTXX$) = "" Then
           '
           If IOBJETX%(i%) > 0 Then
              iobx% = IOBJETX%(i%)
              If PRICAMOBLI% < 1 Then PRICAMOBLI% = i%
              On Error Resume Next
              Text99(iobx%).BackColor = COLOMARCA
              If Err Then
                 On Error GoTo 0
                 GoTo 15
              End If
              '
              ' MUEVE EL FORMULARIO
12            If Text99(iobx%).Top + Text99(iobx%).Height + 500 > Picture2.Height - Picture1.Top Then
                 If VScroll1.Value < VScroll1.Max - 100 Then
                   VScroll1.Value = VScroll1.Value + 100
                   Call VScroll1_Scroll
                   GoTo 12
                 End If
              End If
              '
13            If Text99(iobx%).Top < (-1) * Picture1.Top + 500 Then
                 If VScroll1.Value >= 100 Then
                   VScroll1.Value = VScroll1.Value - 100
                   Call VScroll1_Scroll
                   GoTo 13
                 End If
              End If
              '
15         End If
           '
         End If
       End If
     End If
   Next i%
   '
   Screen.MousePointer = 1
   If PRICAMOBLI% > 0 Then
      Call MENSERR(24, "Campo(s) Obligatorio(s)")
      For ITEX% = 1 To Text99.Count
         ' POSIBLE ERROR 116
         On Error Resume Next
         If Text99(ITEX% - 1).BackColor = COLOMARCA Then
            Text99(ITEX% - 1).BackColor = RGB(255, 255, 255)
         End If
         On Error GoTo 0
      Next ITEX%
      '
      If MODEMO% = 1 Then GoTo 24     ' MODO DE DEMOSTRACION
      '
      Call PONENFOCO(PRICAMOBLI%)
      Exit Sub
   End If
   '
   ' VALIDACION DE RANGO
   For i% = 1 To CANOBJ%
     If TIOBJ$(i%) = "TEXT" Then
       OKX% = VALIRANGO%(i%)
       If OKX% < 1 Then
          '
          If IOBJETX%(i%) > 0 Then
             On Error Resume Next
             Text99(IOBJETX%(i%)).BackColor = COLOMARCA
             On Error GoTo 0
          End If
          Call MENSERR(24, "Valor Fuera de Rango")
          '
          Call PONENFOCO(i%)
          Exit Sub
       End If
       
     End If
   Next i%
   '
24 APRO = "OK"
   If AUTOPAGE% > 0 Then
      If PAGIN <> "PREV" Then
         PAGIN = "NEXT"
      End If
   End If
   '
   Screen.MousePointer = 11
   MODMESS$ = "NOMESS"
   For i% = 1 To CANOBJ%
     If TIOBJ$(i%) = "TEXT" Or TIOBJ$(i%) = "PARAM" Then
       If TASEL$(i%) <> "" Then
         CONDIX$ = TEXREPLA$(CONDSEL$(i%), 1)
         '
         ' verifica si hubo cambios
         VARANTE = VALOBADA(TASEL$(i%), CASEL$(i%), CONDIX$, MODMESS$)
         If VarType(VARANTE) = 7 Then VARANTE = CVINT(VARANTE)         ' CASO FECHA
         If VARANTE <> TEXTO(i%) Then
            If CantRegistros(TASEL$(i%), CONDIX$, MODMESS$) < 1 Then
               ' suprime posibilidad de guardar sin reemplazo de parametros
               If InStr(CONDIX$, "'COMPAR-") > 0 Then GoTo 19
               '
               PPXX% = InStr(CONDIX$, "=")
               CAMPOX$ = TRIM$(Left$(CONDIX$, PPXX% - 1))
               VALOX$ = TRIM$(Mid$(CONDIX$, PPXX% + 1))
               SQLX$ = "INSERT INTO " + TASEL$(i%)
               SQLX$ = SQLX$ + " (" + CAMPOX$ + ") "
               SQLX$ = SQLX$ + " VALUES (" & VALOX$ & ")"
               On Error Resume Next
                FLEXCONN.Execute (FILTSQL$(SQLX$))
               If Err Then
                  On Error GoTo 0
                  If InStr(UCase$(MODMESS$), "NOMESS") < 1 Then
                     Call MENSERR(24, "Imposible Ejecutar '" + SQLX$ + "'.")
                     Call FINAL("")
                  End If
               End If
               On Error GoTo 0
            End If
            Call ACTUREGISTRO(TASEL$(i%), CASEL$(i%), CONDIX$, TEXTO(i%), RAFE&, MODMESS$)
         End If
       End If
     End If
19 Next i%
   '
   If AUTO_PRN$ = "SI" Then
      Call Command4_Click     ' IMPRESION
   End If
   '
   Screen.MousePointer = 1
   On Error Resume Next
   Hide
   On Error GoTo 0
End Sub

Function VALIRANGO%(i%)
   '
   VLR% = 0
   If TRIM$(CStr(TEXTO(i%))) = "" Or TEXTO(i%) = TMASCARA$(FORPA$(i%)) Then
        VLR% = 1: GoTo 99
   End If
   '
   VAMIN$ = TEXREPLA$(VAMINI$(i%))
      If Left$(VAMINI$(i%), 1) = "@" Then VAMIN$ = VAMINI$(i%)
   VAMAX$ = TEXREPLA$(VAMAXI$(i%))
      If Left$(VAMAXI$(i%), 1) = "@" Then VAMAX$ = VAMAXI$(i%)
   FORI$ = UCase$(Left$(FORPA$(i%), 1))
   If FORI$ = "F" Or FORI$ = "G" Or FORI$ = "$" Or FORI$ = "I" Then     ' ES NUMERICO
         NVAMIN = XVALO(VAMIN$)
         NVAMAX = XVALO(VAMAX$)
         NVALAC = XVALO(TEXTO(i%))
         If VAMIN$ <> "" Then
            If NVALAC < NVAMIN Then
               Call MENSERR(24, "Valor Fuera de Rango - Minimo: " & VAMIN$)
               GoTo 99
            End If
         End If
         If VAMAX$ <> "" Then
            If NVALAC > NVAMAX Then
               Call MENSERR(24, "Valor Fuera de Rango - Maximo: " & VAMAX$)
               GoTo 99
            End If
         End If
         VLR% = 1:  GoTo 99
         '
      ElseIf FORI$ = "D" Then                                          ' ES UNA FECHA
         If UCase$(VAMIN$) = "@HOY" Or UCase$(VAMIN$) = "HOY" Then VAMIN$ = Str$(HOY(HOS$))
         If UCase$(VAMAX$) = "@HOY" Or UCase$(VAMAX$) = "HOY" Then VAMAX$ = Str$(HOY(HOS$))
         NVAMIN = XVALO(VAMIN$)
         NVAMAX = XVALO(VAMAX$)
         NVALAC = XVALO(TEXTO(i%))                                     ' PORQUE LAS FECHAS SED GUARDAN COMO NUMERO
         If VAMIN$ <> "" Then
            If NVALAC < NVAMIN Then
               Call MENSERR(24, "Valor Fuera de Rango - Minimo: " & FECHATEX$(NVAMIN))
               GoTo 99
            End If
         End If
         If VAMAX$ <> "" Then
            If NVALAC > NVAMAX Then
               Call MENSERR(24, "Valor Fuera de Rango - Maximo: " & FECHATEX$(NVAMAX))
               GoTo 99
            End If
         End If
         VLR% = 1: GoTo 99
         '
      ElseIf FORI$ = "A" Then                  ' ES UN TEXTO
         VATABLA$ = TRIM$(TEXREPLA$(TABVALI$(i%)))
         If VATABLA$ = "" Then VLR% = 1: GoTo 99
         '
         COMPAXI% = 1
         If Left$(VATABLA$, 3) = "NOT" Then
            VATABLA$ = Mid$(VATABLA$, 4): COMPAXI% = (-1)
         End If
         '
         'IF LEFT$(VATABLA$,1)="#" THEN .... TABLA DAT
         'IF LEFT$(VATABLA$,1)="@" THEN ..... SELECT
         '
         VATABLA$ = TRIM$(UCase$(VATABLA$))
         If VATABLA$ = "" Then VLR% = 1: GoTo 99
         '
15       PPXX% = InStr(VATABLA$, "/")
         PPYY% = InStr(VATABLA$, ";")
         If PPXX% < PPYY% Or PPXX% = 0 Then PPXX% = PPYY%
         If PPXX% < 1 Then PPXX% = 1 + Len(VATABLA$)
         '
         XVALAC$ = TRIM$(UCase$(TEXTO(i%)))
         XVATAB$ = TRIM$(Left$(VATABLA$, PPXX% - 1))
         '
         If TRIM$(XVATAB$) <> "" Then
            If XVALAC$ = XVATAB$ Then
               If COMPAXI% > 0 Then VLR% = 1
               If COMPAXI% < 0 Then VLR% = 0
               GoTo 99
            End If
         End If
         VATABLA$ = TRIM$(Mid$(VATABLA$, PPXX% + 1))
         If VATABLA$ <> "" Then GoTo 15
         '
         ' AQUI CAE SI NO LO ENCONTRO
         If COMPAXI% > 0 Then VLR% = 0
         If COMPAXI% < 0 Then VLR% = 1
         GoTo 99
         '
   End If
   '
   ' AQUI CAE SI NO RECONOCE EL FORMATO
   VLR% = 1
   '
99 VALIRANGO% = VLR%
   If VLR% < 1 Then
      On Error Resume Next
      Text99(i%).SetFocus
      On Error GoTo 0
   End If
   '
End Function
'
Sub PONENFOCO(i%)
    ' PONE EL FOCO EN EL OBJETO DE ORDEN I%
    ' LAS COMENTADAS SON DE LA VERSION ANTERIOR
    ' FALTA CONSIDERAR QUE SEA UN LABEL
    '
    If IOBJETX%(i%) > 0 Then
       On Error Resume Next
       Text99(IOBJETX%(i%)).SetFocus
       On Error GoTo 0
       Exit Sub
    End If
'''           For KKII% = 1 To Text99.Count
'''              If ORITEXT99%(KKII%) = I% Then
'''                 Text99(KKII%).SetFocus
'''                 Exit Sub
'''              End If
'''           Next KKII%
    If IOBJETXM%(i%) > 0 Then
       On Error Resume Next
       Text99M(IOBJETXM%(i%)).SetFocus
       On Error GoTo 0
       Exit Sub
    End If
'''           For KKII% = 1 To Text99M.Count
'''              If ORITEXT99M%(KKII%) = I% Then
'''                 Text99M(KKII%).SetFocus
'''                 Exit Sub
'''              End If
'''           Next KKII%
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   On Error Resume Next
   Label6 = "" 'nombre de adjunto para mail
   STDMAIL.Show 1
   If Err Then
      On Error GoTo 0
      Call COMUNI("Imposible Enviar por Mail\en la Presente Versión.\   \Solicite Actualización de Versión.")
   End If
   On Error GoTo 0
   
'   Call MENSERR(24, "No Disponible en Este Contexto")
   Command6.Enabled = True
End Sub

Private Sub Command99_Click(Index As Integer)
    '
' POSIBLE ERROR 117
    If Text99(Index).Enabled = True Then   ' ESTA CONDICION PARA QUE NO DESPLIEGUE LISTA EN CAMPOS BLOQUEADOS
       ORITE% = ORITEXT99%(Index)
       If ORITE% > 0 Then
          If ORITE% <= CANOBJ% Then
             FOREP$ = FOREPLA$(ORITE%)
             If Left$(FOREP$, 1) = "#" Then
               TASELE$ = TRIM$(Mid$(FOREP$, 2))
               PPXX% = InStr(TASELE$, ">"): If PPXX% > 1 Then TASELE$ = Left$(TASELE$, PPXX% - 1)
               Call SELECHO(TASELE$)
               VDEVU$ = VALACT1$(TASELE$)
               If VDEVU$ <> "" Then
                 Text99(Index) = VDEVU$
                 Call Text99_LostFocus(Index)
               End If
             End If
          End If
       End If
    End If
    '
End Sub

Private Sub Form_Activate()
   If RECARGA% = 1 Then
      RECARGA% = 0
      Call SETNEXTEX(0)
   End If
End Sub

Private Sub Form_DblClick()
  '
  OBUS$ = "LABEL"
  '
2 For i% = CANOBJ% To 1 Step -1
    If TIOBJ$(i%) = OBUS$ Then
      HA = H1(i%): HB = H2(i%)
      If HA > HB Then
         HA = H2(i%): HB = H1(i%)
      End If
      VXA = V1(i%): VXB = V2(i%)
      If VXA > VXB Then
         VXA = V2(i%): VXB = V1(i%)
      End If
      If XMOUSE >= 56.7 * (HA - 0.5) Then
        If XMOUSE <= 56.7 * (HB + 0.5) Then
          If (YMOUSE >= 56.7 * (VXA - 0.5) And YMOUSE <= 56.7 * (VXB + 0.5)) Or (YMOUSE >= 56.7 * (VXB - 0.5) And YMOUSE <= 56.7 * (VXA + 0.5)) Then
               GoTo 6
          End If
        End If
      End If
    End If
  Next i%
  '
  If OBUS$ = "LABEL" Then OBUS$ = "PARAM": GoTo 2
  If OBUS$ = "PARAM" Then OBUS$ = "TEXT": GoTo 2
  If OBUS$ = "TEXT" Then OBUS$ = "LINE": GoTo 2
  If OBUS$ = "LINE" Then OBUS$ = "GRAPH": GoTo 2
  If OBUS$ = "GRAPH" Then OBUS$ = "GRID": GoTo 2
  If OBUS$ = "GRID" Then OBUS$ = "BOX": GoTo 2
  Exit Sub
  '
6 ' PRESENTA DATOS DEL OBJETO
  If TIOBJ$(i%) = "GRAPH" Then
      TTXZ$ = TRIM$(UCase$(TEXTO(i%)))
      If TTXZ$ = "@IMAGEN" Or TTXZ$ = "@FOTO" Or TTXZ$ = "@PICTURE" Then
         ARIMAG$ = GETPICTURE$                    ' ESTO INVOCA LA FUNCION QUE TOMA LA IMAGEN
         TEXTO(i%) = ARIMAG$
         If EXISTE%(ARIMAG$) > 0 Then
            Call DIBUIMAG(STDSCREEN.Picture1, ARIMAG$, H1(i%), H2(i%), V1(i%), V2(i%))
         End If
      End If
  End If
  '
  If ARDOCLI$(i%) <> "" Then
     If HABILAPLI%("SHELLEXEC") >= 0 Then
        Call ShellExecute(Me.hwnd, "Open", ARDOCLI$(i%), "", "", 1)
     End If
  End If
  '
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  '
  If KeyCode = 33 Then
      Call PAGANT1_Click
    ElseIf KeyCode = 34 Then
      Call PAGSIG1_Click
    ElseIf KeyCode = 40 Then
      VS = VScroll1.Value + VScroll1.Max / 10
      If VS > VScroll1.Max Then VS = VScroll1.Max
      VScroll1.Value = VS
      Call VScroll1_Scroll
    ElseIf KeyCode = 38 Then
      VS = VScroll1.Value - VScroll1.Max / 10
      If VS < VScroll1.Min Then VS = VScroll1.Min
      VScroll1.Value = VS
      Call VScroll1_Scroll
    ElseIf KeyCode = 39 Then
      HS = HScroll1.Value + HScroll1.Max / 10
      If HS > HScroll1.Max Then HS = HScroll1.Max
      HScroll1.Value = HS
      Call HScroll1_Scroll
    ElseIf KeyCode = 37 Then
      HS = HScroll1.Value - HScroll1.Max / 10
      If HS < HScroll1.Min Then HS = HScroll1.Min
      HScroll1.Value = HS
      Call HScroll1_Scroll
  End If
  '
End Sub

Private Sub Form_Load()
    PINHOR = Picture1.Left
    PINVER = Picture1.Top
    'VALIDA LA EXISTENCIA DEL FORMULARIO ENTRAPASS17 EN EL PROYECTO PORQUE SINO SE PRODUCE UN ERROR 424
    On Error Resume Next
    A1$ = ENTRAPASS17.Text1
    If Err Then
        Call MENSERR(24, "Imposible Presentar Formulario Standard Screen. Falta Formulario ENTRAPASS17.\Favor de Consultar con Mesa de Ayuda")
        On Error GoTo 0
        Unload Me
    End If
    On Error GoTo 0
    '
    Call APLICACIONSKINS(Me)
    RECARGA% = 1
    Combo1.Clear
    Dim PRX As Printer
    For Each PRX In Printers
      Combo1.AddItem PRX.DeviceName
    Next
    On Error Resume Next
    Combo1.TEXT = Printer.DeviceName
    On Error GoTo 0
    '
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
      '
      OBUS$ = "LABEL"
2     For i% = CANOBJ% To 1 Step -1
        If TIOBJ$(i%) = OBUS$ Then
          HA = H1(i%): HB = H2(i%)
          If HA > HB Then
             HA = H2(i%): HB = H1(i%)
          End If
          VXA = V1(i%): VXB = V2(i%)
          If VXA > VXB Then
             VXA = V2(i%): VXB = V1(i%)
          End If
          If XMOUSE >= 56.7 * (HA - 0.5) Then
            If XMOUSE <= 56.7 * (HB + 0.5) Then
              If (YMOUSE >= 56.7 * (VXA - 0.5) And YMOUSE <= 56.7 * (VXB + 0.5)) Or (YMOUSE >= 56.7 * (VXB - 0.5) And YMOUSE <= 56.7 * (VXA + 0.5)) Then
                   GoTo 6
              End If
            End If
          End If
        End If
      Next i%
      '
      If OBUS$ = "LABEL" Then OBUS$ = "PARAM": GoTo 2
      If OBUS$ = "PARAM" Then OBUS$ = "TEXT": GoTo 2
      If OBUS$ = "TEXT" Then OBUS$ = "LINE": GoTo 2
      If OBUS$ = "LINE" Then OBUS$ = "GRAPH": GoTo 2
      If OBUS$ = "GRAPH" Then OBUS$ = "GRID": GoTo 2
      If OBUS$ = "GRID" Then OBUS$ = "BOX": GoTo 2
      Exit Sub
      '
6     ' PRESENTA DATOS DEL OBJETO
      On Error Resume Next
      DEVALOBA17.Label14 = i%
      If Err Then
         On Error GoTo 0
         Call MENSERR(24, "Falta Formulario 'DEVALOBA17' en Proyecto '" + App.EXEName + "'")
         Exit Sub
      End If
      '
      On Error GoTo 0
      DEVALOBA17.Caption = "Propiedades de: " + ONAME$(i%) + " (" + OBUS$ + ")"
      DEVALOBA17.Command1.Enabled = False
      DEVALOBA17.Command2.Enabled = False
      DEVALOBA17.Command3.Enabled = False
      DEVALOBA17.Command4.Enabled = False
      DEVALOBA17.Command5.Enabled = False
      '
      Line2.X1 = 0: Line2.Y1 = 0
      Line2.X2 = 56.7 * H1(i%): Line2.Y2 = 56.7 * V1(i%)
      Line2.Visible = True
      DEVALOBA17.Show 1
      Line2.Visible = False
      '
    End If
End Sub

Private Sub label1_Change()
   '
   Picture1.Cls
   ' POSIBLE ERROR 8 - CAPTURADO
   On Error Resume Next
   NUPIC% = Label99.Count
   For KKI% = 1 To NUPIC% - 1: Label99(KKI%).Caption = "": Label99(KKI%).Visible = False: Next KKI%
   NUPIC% = Text99.Count
   For KKI% = 1 To NUPIC% - 1: Text99(KKI%).TEXT = "": Text99(KKI%).Visible = False: Next KKI%
   NUPIC% = Command99.Count
   For KKI% = 1 To NUPIC% - 1:  Command99(KKI%).Visible = False: Next KKI%
   NUPIC% = Text99M.Count
   For KKI% = 1 To NUPIC% - 1: Text99M(KKI%).TEXT = "": Text99M(KKI%).Visible = False: Next KKI%
   NUPIC% = GRID99.Count
   For KKI% = 1 To NUPIC% - 1: GRID99(KKI%).Clear: GRID99(KKI%).Visible = False: Next KKI%
   On Error GoTo 0
   APRO = "": PAGIN = ""
   '
   APLITO$ = Label1
   If APLITO$ <> "" Then
      FILTX$ = "": CAPARINV% = 0
      PPXX% = InStr(APLITO$, "(")
      If PPXX% > 0 Then
         PPYY% = InStr(PPXX% + 1, APLITO$, ")"): If PPYY% < PPXX% Then PPYY% = 1 + Len(APLITO$)
         FILTX$ = Mid$(APLITO$, PPXX% + 1, PPYY% - PPXX% - 1)
         APLITO$ = Left$(APLITO$, PPXX% - 1)
      End If
      '
      Screen.MousePointer = 11
      CANFI% = 0
      Do While Len(FILTX$) > 0
         PPXX% = InStr(FILTX$, ";"): If PPXX% < 1 Then PPXX% = 1 + Len(FILTX$)
         TSEL$ = TRIM$(Left$(FILTX$, PPXX% - 1))
         FILTX$ = Mid$(FILTX$, PPXX% + 1)
         If TSEL$ <> "" Then
            CANFI% = CANFI% + 1
            ONAME$(CANFI%) = "COMPAR-" & CStr(CANFI% - 1)
            TEXTO(CANFI%) = TSEL$
            ARDOCLI$(CANFI%) = ""
            '
            TIOBJ$(CANFI%) = "COMPAR"
            H1(CANFI%) = 0
            V1(CANFI%) = 0
            H2(CANFI%) = 0
            V2(CANFI%) = 0
            CONMARCA%(CANFI%) = 0
            '
            FUENTE$(CANFI%) = ""
            TAFUEN$(CANFI%) = ""
            FORPA$(CANFI%) = ""
            '
            NEGRI(CANFI%) = 0
            ITALI(CANFI%) = 0
            SUBRA(CANFI%) = 0
            INVIS(CANFI%) = 0
            ENCAS(CANFI%) = 0
            '
            FOREPLA$(CANFI%) = ""
            TASEL$(CANFI%) = ""
            CASEL$(CANFI%) = ""
            CONDSEL$(CANFI%) = ""
            OBLIGA%(CANFI%) = 1
            LOCKFI$(CANFI%) = ""
         End If
      Loop
      FORNA$ = TRIM$(APLITO$)
      Call CARGADATFOR(FORNA$, CANFI%)
      Call MUESTRAFOR
      Screen.MousePointer = 1
   End If
   '
End Sub

Private Sub Label99_Change(Index As Integer)
   ORITE% = ORILAB99%(Index)
   If ORITE% > 0 Then
      FOREP$ = FOREPLA$(ORITE%)
      If Left$(FOREP$, 1) = "#" Then
         FOREP$ = Mid$(FOREP$, 2)
         PPXX% = InStr(FOREP$, ">")
         If PPXX% > 1 Then
            FORMECO$ = TRIM$(Mid$(FOREP$, PPXX% + 1))
            FOREP$ = Left$(FOREP$, PPXX% - 1)
 ' POSIBLE ERROR 9
           TECOL99$ = FORMATNUM$(ECOARCH$(FOREP$, Label99(Index).Caption), FORMECO$)
            '
            Picture1.CurrentY = STDSCREEN.Label99(Index).Top
            Picture1.CurrentX = STDSCREEN.Label99(Index).Left + STDSCREEN.Label99(Index).Width + 100
            Picture1.Font = FUENTE$(ORITE%)
            Picture1.FontSize = TAFUEN(ORITE%)
            Picture1.FontBold = NEGRI(ORITE%)
            Picture1.FontItalic = ITALI(ORITE%)
            Picture1.ForeColor = RGB(0, 0, 0)
            Picture1.Print TECOL99$
            '
         End If
      End If
   End If

End Sub

Private Sub Label99_DblClick(Index As Integer)
    Call Form_DblClick
End Sub

Private Sub Label99_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Form_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Label99_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
' POSIBLE ERROR 10
   XMOUSE = X + Label99(Index).Left: YMOUSE = Y + Label99(Index).Top
   Label2 = FORMATNUM(XMOUSE / 56.7, "F6.1")
   Label3 = FORMATNUM(YMOUSE / 56.7, "F6.1")
End Sub

Private Sub PAGANT_Click()
   Call Command5_Click
   If APRO = "OK" Then PAGIN = "PREV"
End Sub

Private Sub PAGANT1_Click()
   Call PAGANT_Click
End Sub

Private Sub PAGSIG_Click()
   Call Command5_Click
   If APRO = "OK" Then PAGIN = "NEXT"
End Sub

Private Sub PAGSIG1_Click()
   Call PAGSIG_Click
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   XMOUSE = X: YMOUSE = Y
   If Button = 2 Then
      Call Form_MouseDown(Button, Shift, X, Y)
   End If
End Sub

Private Sub Picture1_DBLClick()
   Call Form_DblClick
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   XMOUSE = X: YMOUSE = Y
   Label2 = FORMATNUM(XMOUSE / 56.7, "F6.1")
   Label3 = FORMATNUM(YMOUSE / 56.7, "F6.1")
End Sub

Private Sub Picture2_DblClick()
   Call Form_DblClick
End Sub

Private Sub Text99_Change(Index As Integer)
   ORITE% = ORITEXT99%(Index)
   If ORITE% > 0 Then
      CARMAN%(ORITE%) = 1
      FOREP$ = FOREPLA$(ORITE%)
      If Left$(FOREP$, 1) = "#" Then
         FOREP$ = Mid$(FOREP$, 2)
         PPXX% = InStr(FOREP$, ">")
         If PPXX% > 1 Then
            FORMECO$ = TRIM$(Mid$(FOREP$, PPXX% + 1))
            FOREP$ = Left$(FOREP$, PPXX% - 1)
            '
            If FOREP$ <> "" Then
               TPF% = TIPFIELD%(FOREP$, 1)
               ASELE$ = Text99(Index).TEXT
               If TPF% <= 2 Then
                   If LENFIELD%(FOREP$, 1) < 2 Then
                       ASELE$ = Chr$(XVALO(ASELE$))
                     Else
                       ASELE$ = MKI$(XVALO(ASELE$))
                   End If
                 ElseIf TPF% <= 4 Then
                   If LENFIELD%(FOREP$, 1) < 8 Then
                       ASELE$ = MKS$(XVALO(ASELE$))
                     Else
                       ASELE$ = MKD$(XVALO(ASELE$))
                   End If
                 ElseIf TPF% = 5 Then
                   ASELE$ = MKI$(FECHANUM(ASELE$))
               End If
               '
               ' ECO99(Index) = FORMATNUM$(ECOARCH$(FOREP$, Text99(Index).TEXT), FORMECO$)
               ECO99(Index) = FORMATNUM$(ECOARCH$(FOREP$, ASELE$), FORMECO$)
            End If
         End If
      End If
   End If
End Sub

Private Sub Text99_Click(Index As Integer)
    ORITE% = ORITEXT99%(Index)
    If ORITE% > 0 Then
      If ORITE% <= CANOBJ% Then
        FOREP$ = FOREPLA$(ORITE%)
        If Left$(UCase$(FOREP$), 9) = "@PASSWORD" Then
            On Error Resume Next
            If ENTRAPASS17.Visible = False Then
              On Error GoTo 0
              Call Text99_GotFocus(Index)
            End If
            On Error GoTo 0
            Exit Sub
          Else
            If LOCKFI$(ORITE%) <> "" Then
              If FOREP$ <> "" Then
                Call SETNEXTEX(ORITE%)
              End If
            End If
        End If
      End If
    End If
End Sub

Private Sub Text99_DblClick(Index As Integer)
    Call Form_DblClick
    ORITE% = ORITEXT99%(Index)
    If ORITE% > 0 Then
       If ORITE% <= CANOBJ% Then
          FOREP$ = FOREPLA$(ORITE%)
          '
          If LOCKFI$(ORITE%) <> "" Then
            If Left$(FOREP$, 1) = "@" Then
              Call SETNEXTEX(ORITE%)
              Exit Sub
            End If
          End If
          '
          If UCase$(FOREP$) = "@USERNAME" Then
              TREP = USERNAME$: GoTo 90
            ElseIf UCase$(FOREP$) = "@AHORA" Then
              TREP = Time: GoTo 90
            ElseIf UCase$(FOREP$) = "@HOY" Then
              TREP = FECHATEX$(HOY(HOS$)): GoTo 90
          End If
          '
          If InStr(FOREP$, "/") > 0 Then
' POSIBLE ERROR 12
             TREP = CONMUTEX(Text99(Index).TEXT, FOREP$)
             GoTo 90
          End If
          '
       End If
    End If
    Exit Sub
    '
90  Text99(Index) = FORMATNUM(TREP, FORPA$(ORITE%))
    Call Text99_LostFocus(Index)
    '
End Sub

Private Sub Text99_GotFocus(Index As Integer)
    '
    Static UTIM, UINDEX
    '
    If Index = UINDEX Then
       If Timer < UTIM + 0.2 Then
          Exit Sub
       End If
    End If
    '
    Text99(Index).BackColor = RGB(255, 220, 220)
    Text99(Index).SelStart = 0
    Text99(Index).SelLength = Len(RTrim$(Text99(Index).TEXT))
    '
    Call MUEVEFORM(Index)        ' MUEVE EL FORMULARIO
    '
    ORITE% = ORITEXT99%(Index)
    If ORITE% > 0 And ORITE% <= CANOBJ% Then
      FOREP$ = FOREPLA$(ORITE%)
      ACNAM$ = UCase$(TRIM$(ONAME$(ORITE%)))
      '
      ' AQUI EL CODIGO CORRESPONDIENTE A @PASSWORD
      If Left$(UCase$(FOREP$), 9) = "@PASSWORD" Then
          '
          Text99(Index) = ""
          If EXISTENTRAPASS% < 1 Then GoTo 99 ' VERIFICA EXISTENCIA DEL FORM
          '
          ' VERIFICA VALIDACION DE DATOS EN TODOS LOS CAMPOS A BLOQUEAR
          For KI% = 1 To Text99.Count
            ORITE1% = ORITEXT99%(KI%)
            If ORITE1% > 0 And ORITE1% <= CANOBJ% Then
              LOCKFI1$ = UCase$(TRIM$(LOCKFI$(ORITE1%)))
              If LOCKFI1$ <> "" Then
                 If LOCKFI1$ = ACNAM$ Then
                   If LOCKFI1$ = ONAME$(ORITE1%) Then GoTo 29
                   FOREP1$ = UCase$(FOREPLA$(ORITE1%))
                   If Left$(UCase$(FOREP1$), 9) = "@PASSWORD" Then GoTo 29
                   If UCase$(FOREP1$) = "@USERNAME" Then GoTo 29
                   If UCase$(FOREP1$) = "@AHORA" Then GoTo 29
                   If UCase$(FOREP1$) = "@HOY" Then GoTo 29
                   If InStr(FOREP1$, "{" + LOCKFI1$ + "}") > 0 Then GoTo 29 ' PARA QUE LOS CAMPOS ECO DEL PASSWORD PASEN
                   '
                   ' VLDX% = VALIVATE%(KI%)
                   VLDX% = VALIVATE%(KI%)
                   If VLDX% < 1 Then
                       On Error Resume Next
                       Text99(Index).BackColor = RGB(255, 255, 255)
                       On Error GoTo 0
                       Exit Sub
                    End If
                 End If
              End If
            End If
29        Next KI%
          '
          Unload ENTRAPASS17
          ENTRAPASS17.Text2 = ""
          ENTRAPASS17.Left = Text99(Index).Left + Picture1.Left + Me.Left
          ENTRAPASS17.Top = Me.Top + Picture1.Top + Text99(Index).Top - ENTRAPASS17.Height - 100
          If ENTRAPASS17.Top < 1000 Then
             ENTRAPASS17.Top = Me.Top + Picture2.Top + Picture1.Top + Text99(Index).Top + 1000
          End If
          BANDENTRAPASS% = 1
          '
          For KKII% = 1 To Text99.Count
             Text99(KKII% - 1).BackColor = RGB(255, 255, 255)
          Next KKII%
          '
          Text99(Index).BackColor = RGB(100, 255, 100)
          Text99(Index) = ""
          
          Call MUEVEFORM(Index)      ' MUEVE EL FORMULARIO
          INDEXTEX% = Index
          ENTRAPASS17.Show 1
          If ENTRAPASS17.Text1 = "" Then
                UTIM = Timer: UINDEX = Index             ' CANCELO
              ElseIf ENTRAPASS17.Text1 <> "*" Then
                UTIM = Timer: UINDEX = Index             ' INGRESO VALIDO
          End If
          BANDENTRAPASS% = 0
          '
             '
          End If
'        End If
' POSIBLE ERROR 17
        Text99(Index).BackColor = RGB(255, 255, 255)
      End If
      '
99    If LOCKFI$(ORITE%) <> "" Then
        If FOREP$ <> "" Then
          Call SETNEXTEX(ORITE%)
        End If
      End If
      '
    ' End If
    '
End Sub

Sub MUEVEFORM(INDEXXX)
    '
12  If Text99(INDEXXX).Top + Text99(INDEXXX).Height + 500 > Picture2.Height - Picture1.Top Then
       If VScroll1.Value < VScroll1.Max - 100 Then
         VScroll1.Value = VScroll1.Value + 100
         Call VScroll1_Scroll
         GoTo 12
       End If
    End If
    '
13  If Text99(INDEXXX).Top < (-1) * Picture1.Top + 500 Then
       If VScroll1.Value >= 100 Then
         VScroll1.Value = VScroll1.Value - 100
         Call VScroll1_Scroll
         GoTo 13
       End If
    End If
    '
End Sub
'
Function EXISTENTRAPASS%()
    '
    EXISXX% = 1
    On Error Resume Next
    ENTRAPASS17.Text1 = ""
    If Err Then
        On Error GoTo 0
        Call MENSERR(24, "Imposible Registro de Firma Electronica.\   \Falta Formulario 'ENTRAPASS17.FRM' en Proyecto '" + App.EXEName + "'.")
        EXISXX% = 0
        GoTo 99
    End If
    On Error GoTo 0
    '
99  EXISTENTRAPASS% = EXISXX%
    '
End Function
'
Function VALIVATE%(KI%)
    '
    Static UFORMU$, KINOVA$
    '
    COLOMARCA = RGB(255, 100, 100)
    ORITE1% = ORITEXT99%(KI%)
    '
    If OBLIGA%(ORITE1%) > 0 Then
       TTXX$ = TEXTO(ORITE1%)
       If TRIM$(TTXX$) = "0" Then TTXX$ = TRIM$(Text99(KI%))
       If TRIM$(TTXX$) = "" Then
         '
         If IOBJETX%(ORITE1%) > 0 Then
            iobx% = KI%    '   IOBJETX%(ORITE1%)
            On Error Resume Next
            Text99(iobx%).BackColor = COLOMARCA
            '
            If Err Then
               On Error GoTo 0
               GoTo 15
            End If
            '
            ' MUEVE EL FORMULARIO
12          If Text99(iobx%).Top + Text99(iobx%).Height + 500 > Picture2.Height - Picture1.Top Then
               If VScroll1.Value < VScroll1.Max - 100 Then
                 VScroll1.Value = VScroll1.Value + 100
                 Call VScroll1_Scroll
                 GoTo 12
               End If
            End If
            '
13          If Text99(iobx%).Top < (-1) * Picture1.Top + 500 Then
               If VScroll1.Value >= 100 Then
                 VScroll1.Value = VScroll1.Value - 100
                 Call VScroll1_Scroll
                 GoTo 13
               End If
            End If
            '
15       End If
         '
         Call MENSERR(24, "Imposible Registro de Firma Electronica.\   \Campo Obligatorio.")
         VALIVATE% = 0
         GoTo 99
         '
       Else
         FORI$ = FORPA$(ORITE1%)
         If Left$(FORI$, 1) = "F" Or Left$(FORI$, 1) = "G" Or Left$(FORI$, 1) = "$" Or Left$(FORI$, 1) = "I" Then
            If IsNumeric(Text99(KI%)) = False Then
               Call MENSERR(24, "Ingreso No Numérico")
               VALIVATE% = 0
               GoTo 99
            End If
         End If
       End If
    End If
    '
    If VALIRANGO%(ORITE1%) < 1 Then
        VALIVATE% = 0
        GoTo 99
    End If
    '
98  VALIVATE% = 1
    '
99  'If VALIVATE% < 1 Then
    '   On Error Resume Next
    '   Text99(KI%).SetFocus
    '   On Error GoTo 0
    'End If
    '
End Function

Private Sub Text99_KeyPress(Index As Integer, KeyAscii As Integer)
   ORITE% = ORITEXT99%(Index)
   If ORITE% > 0 And ORITE% <= CANOBJ% Then
      FOREP$ = FOREPLA$(ORITE%)
      If Left$(UCase$(FOREP$), 9) = "@PASSWORD" Then
         KeyAscii = 0: Text99(Index) = ""
         Call Text99_GotFocus(Index)
         Exit Sub
      End If
      '
      If KeyAscii% = 13 Then
         Call Text99_LostFocus(Index)
         Call SETNEXTEX(ORITE%)
      End If
   End If
End Sub

Private Sub Text99_LostFocus(Index As Integer)
    COLOAN = Text99(Index).BackColor
    If ENTRAPASS17.Visible = False Then Text99(Index).BackColor = RGB(255, 255, 255)
    ORITE% = ORITEXT99%(Index)
    If ORITE% > 0 Then
      If ORITE% <= CANOBJ% Then
        If BANDENTRAPASS% = 1 Then
           Text99(Index).BackColor = RGB(100, 255, 100)
        End If
        '
        If TRIM$(Text99(Index)) = "" Then
           FORIXX$ = TRIM$(Left$(UCase$(FORPA$(ORITE%)), 1))
           If FORIXX$ = "F" Or FORIXX$ = "G" Or FORIXX$ = "$" Or FORIXX$ = "I" Then
                VALOR = 0
              Else
                VALOR = ""
           End If
           GoTo 25
        End If
        '
        If TRIM$(Text99(Index)) <> "" Then
          FORI$ = FORPA$(ORITE%)
          If Left$(FORI$, 1) = "F" Or Left$(FORI$, 1) = "G" Or Left$(FORI$, 1) = "$" Or Left$(FORI$, 1) = "I" Then
              If IsNumeric(Text99(Index)) = False Then
                Call MENSERR(24, "Ingreso No Numérico")
                On Error Resume Next
                Text99(Index).SetFocus
                On Error GoTo 0
                Exit Sub
              End If
              VALOR = XVALO(Text99(Index))
              '
            ElseIf Left$(FORI$, 1) = "D" Then
              FEX = FECHANUM(TRIM$(Text99(Index)))
              If FEX < 1 Then
                Call MENSERR(24, "Corresponde Ingreso Fecha")
                On Error Resume Next
                Text99(Index).SetFocus
                On Error GoTo 0
                Exit Sub
              End If
              VALOR = FEX
              '
            ElseIf Left$(FORI$, 1) = "H" Or Left$(FORI$, 1) = "T" Then
              VVXX$ = REPLACAR$(Text99(Index), ".", ":")
              VVXX$ = REPLACAR$(VVXX$, ",", ":")
              VVXX$ = REPLACAR$(VVXX$, "/", ":")
              VVXX$ = REPLACAR$(VVXX$, ";", ":")
              VALOR = VVXX$
            Else
              VALOR = TRIM$(Text99(Index))
          End If
          '
25        TEXTO(ORITE%) = VALOR
          If Left$(UCase$(FOREPLA$(ORITE%)), 9) = "@PASSWORD" Then
             GoTo 19 ' LOS PASSWORD SE VALIDAN AL GOTFOCUS, NO AL LOSTFOCUS
          End If
          '
          If VALIRANGO%(ORITE%) < 1 Then
              On Error Resume Next
              Text99(Index).BackColor = RGB(255, 100, 100)
              On Error GoTo 0
              If Left$(UCase$(FOREPLA$(ORITE%)), 9) = "@PASSWORD" Then
                   Call MENSERR(24, "Usuario / Clave no Válido\para la Firma Electrónica Requerida.")
                   On Error Resume Next
                   Text99(Index).SetFocus
                   On Error GoTo 0
                 Else
                   ' Call MENSERR(24, "Ingreso no Válido")
                   ' Call SETNEXTEX(ORITE%)
              End If
              Exit Sub
          End If
          '
19        CARMAN%(ORITE%) = 1
          Text99(Index) = FORMATNUM$(VALOR, FORI$)
          Call RECALCUTO(ORITE%)
        End If
      End If
      '
      On Error Resume Next
      If Left$(UCase$(FOREPLA$(ORITE%)), 9) <> "@PASSWORD" Then Text99(Index).BackColor = RGB(255, 255, 255)
      On Error GoTo 0
      '
    End If
    '
End Sub

Sub RECALCUTO(CACTI%)
    '
    For i% = 1 To CANOBJ%
      If i% <> CACTI% Then
          If TIOBJ$(i%) = "PARAM" Or TIOBJ$(i%) = "TEXT" Then
            If TIOBJ$(i%) <> "TEXT" Or i% > CACTI% Then
              TEXAN = TEXTO(i%)
              Call CARGATEXTO(i%, TOOLTI$)
              iobx% = IOBJETX%(i%)
              On Error Resume Next
              If iobx% > 0 Then
                 If Left$(UCase$(App.EXEName), 6) = "DSGFRM" Then
                    ' POSIBLE ERROR 1118 - CAPTURADO
                    On Error Resume Next
                    Text99(iobx%).ToolTipText = TOOLTI$
                    On Error GoTo 0
                 End If
              End If
              On Error GoTo 0
            End If
          End If
      End If
    Next i%
    '
    For J% = 1 To Text99.Count - 1
      If Text99(J%).Enabled = True Then
        If MULIN(J%) < 1 Then
          ORITE% = ORITEXT99(J%)
          FOREP$ = FOREPLA$(ORITE%)
          If Left$(UCase$(FOREP$), 9) = "@PASSWORD" Then GoTo 29
          If UCase$(FOREP$) = "@AHORA" Then GoTo 29
          If UCase$(FOREP$) = "@HOY" Then GoTo 29
          '
          If TEXTO(ORITE%) = TMASCARA$(FORPA$(ORITE%)) Then GoTo 29
          If InStr(TEXTO(J%), "/") > 0 Then GoTo 29
          '
          ' POSIBLE ERROR 1119
          On Error Resume Next
          Text99(J%) = FORMATNUM$(TEXTO(ORITE%), FORPA$(ORITE%))
          On Error GoTo 0
        End If
      End If
29  Next J%
    '
    For J% = 1 To Label99.Count - 1
       ORITE% = ORILAB99(J%)
       If MULIN(ORITE%) < 1 Then
         Label99(J%) = FORMATNUM$(TEXTO(ORITE%), FORPA$(ORITE%))
       End If
    Next J%
    '
End Sub

Private Sub Text99_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Form_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Text99_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   XMOUSE = X + Text99(Index).Left: YMOUSE = Y + Text99(Index).Top
   Label2 = FORMATNUM(XMOUSE / 56.7, "F6.1")
   Label3 = FORMATNUM(YMOUSE / 56.7, "F6.1")
End Sub

Private Sub Text99M_Change(Index As Integer)
   ORITE% = ORITEXT99M%(Index)
   If ORITE% > 0 Then
      CARMAN%(ORITE%) = 1
   End If
End Sub

Private Sub Text99M_DblClick(Index As Integer)
    Call Form_DblClick
End Sub

Private Sub Text99M_GotFocus(Index As Integer)
    Text99M(Index).BackColor = RGB(255, 220, 220)
End Sub

Private Sub Text99M_LostFocus(Index As Integer)
    If TRIM$(Text99M(Index)) <> "" Then
      ORITE% = ORITEXT99M%(Index)
      If ORITE% > 0 Then
        If ORITE% <= CANOBJ% Then
          TEXTO(ORITE%) = Text99M(Index)
          FORI$ = FORPA$(ORITE%)
          If Left$(FORI$, 1) = "F" Or Left$(FORI$, 1) = "G" Or Left$(FORI$, 1) = "$" Or Left$(FORI$, 1) = "I" Then
              If IsNumeric(Text99M(Index)) = False Then
                Call MENSERR(24, "Ingreso No Numérico")
                On Error Resume Next
                Text99M(Index).SetFocus
                On Error GoTo 0
                Exit Sub
              End If
              VALOR = XVALO(Text99M(Index))
              '
            ElseIf Left$(FORI$, 1) = "D" Then
              FEX = FECHANUM(TRIM$(Text99M(Index)))
              If FEX < 1 Then
                Call MENSERR(24, "Corresponde Ingreso Fecha")
                On Error Resume Next
                Text99M(Index).SetFocus
                On Error GoTo 0
                Exit Sub
              End If
              VALOR = FEX
              '
            Else
              VALOR = Text99M(Index)
          End If
          '
          If MULIN(ORITE%) <> 1 Then
              Text99M(Index) = FORMATNUM$(VALOR, FORI$)
          End If
          'GoSub RECALCUTO
          Call RECALCUTO(ORITE%)
        End If
      End If
    End If
    Text99M(Index).BackColor = RGB(255, 255, 255)
    '
End Sub

Sub SETNEXTEX(NOBJ%)
    '
    INIBU% = NOBJ%
    '
    For KKII% = 1 To Text99.Count
       On Error Resume Next
       Text99(KKII% - 1).BackColor = RGB(255, 255, 255)
       On Error GoTo 0
    Next KKII%
    '
Exit Sub
'''    If NOBJ% > 0 Then
'''       If FORMALTER.Visible = False Then
'''          MsgBox "Pulse 'Aceptar' para Habilitar Ingreso de Dato / Firma Electronica"
'''       End If
'''    End If
    '
15  PROTEX% = INIBU%
17  PROTEX% = PROTEX% + 1
    If PROTEX% > CANOBJ% Then
       If INIBU% > 0 Then
           INIBU% = 0
           GoTo 15
         Else
           Exit Sub
       End If
    End If
    If TIOBJ$(PROTEX%) <> "TEXT" Then GoTo 17
    If INVIS(PROTEX%) = 1 Then GoTo 17
    If LOCKFI$(PROTEX%) <> "" Then
       If FOREPLA$(PROTEX%) <> "" Then
          If UCase$(FOREPLA$(PROTEX%)) <> "@PASSWORD" Then
             GoTo 17
          End If
       End If
    End If
    '
    On Error GoTo 30
    If IOBJETX%(PROTEX%) > 0 Then
         Text99(IOBJETX%(PROTEX%)).SetFocus
         Call Text99_GotFocus(IOBJETX%(PROTEX%))
       ElseIf IOBJETXM%(PROTEX%) > 0 Then
         Text99M(IOBJETXM%(PROTEX%)).SetFocus
         Call Text99M_GotFocus(IOBJETX%(PROTEX%))
    End If
26  On Error GoTo 0
    Exit Sub
    '
30  Resume 31
31  On Error GoTo 0
    GoTo 17
    '
End Sub

Private Sub Text99M_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Form_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Text99M_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   XMOUSE = X + Text99M(Index).Left: YMOUSE = Y + Text99M(Index).Top
   Label2 = FORMATNUM(XMOUSE / 56.7, "F6.1")
   Label3 = FORMATNUM(YMOUSE / 56.7, "F6.1")
End Sub
'
Sub VScroll1_Scroll()
   Picture1.Top = -30 - VScroll1.Value + PINVER
   POVERTI = VScroll1.Value
End Sub
'
Private Sub HScroll1_Scroll()
   Picture1.Left = -30 - HScroll1.Value + PINHOR
End Sub

Private Sub Text99_GotFocus_VIE(Index As Integer)
    '
    Static TANTE, INDEXAN
    '
    If Index = INDEXAN Then Exit Sub
    INDEXAN = Index
    '
' POSIBLE ERROR 13
    Text99(Index).BackColor = RGB(255, 220, 220)
    Text99(Index).SelStart = 0
    Text99(Index).SelLength = Len(RTrim$(Text99(Index).TEXT))
    '
            ' MUEVE EL FORMULARIO
12          If Text99(Index).Top + Text99(Index).Height + 500 > Picture2.Height - Picture1.Top Then
               If VScroll1.Value < VScroll1.Max - 100 Then
                 VScroll1.Value = VScroll1.Value + 100
                 Call VScroll1_Scroll
                 GoTo 12
               End If
            End If
            '
13          If Text99(Index).Top < (-1) * Picture1.Top + 500 Then
               If VScroll1.Value >= 100 Then
                 VScroll1.Value = VScroll1.Value - 100
                 Call VScroll1_Scroll
                 GoTo 13
               End If
            End If
            '
    ORITE% = ORITEXT99%(Index)
    If ORITE% > 0 And ORITE% <= CANOBJ% Then
      FOREP$ = FOREPLA$(ORITE%)
      ACNAM$ = UCase$(TRIM$(ONAME$(ORITE%)))
      If Left$(UCase$(FOREP$), 9) = "@PASSWORD" Then
'        If Timer >= TANTE + 1 Then
          '
          Text99(Index) = ""
          '
          ' VERIFICA EXISTENCIA DEL FORM
          On Error Resume Next
          ENTRAPASS17.Text1 = ""
          If Err Then
             On Error GoTo 0
             Call MENSERR(24, "Imposible Registro de Firma Electronica.\   \Falta Formulario 'ENTRAPASS17.FRM' en Proyecto '" + App.EXEName + "'.")
             GoTo 99
          End If
          On Error GoTo 0
          '
          '
          ' VERIFICA VALIDACION DE DATOS EN TODOS LOS CAMPOS A BLOQUEAR
          For KI% = 1 To Text99.Count
            ORITE1% = ORITEXT99%(KI%)
            If ORITE1% > 0 And ORITE1% <= CANOBJ% Then
              LOCKFI1$ = UCase$(TRIM$(LOCKFI$(ORITE1%)))
              If LOCKFI1$ <> "" Then
                 If LOCKFI1$ = ACNAM$ Then
                   If LOCKFI1$ = ONAME$(ORITE1%) Then GoTo 29
                   FOREP1$ = UCase$(FOREPLA$(ORITE1%))
                   If Left$(UCase$(FOREP1$), 9) = "@PASSWORD" Then GoTo 29
                   If UCase$(FOREP1$) = "@USERNAME" Then GoTo 29
                   If UCase$(FOREP1$) = "@AHORA" Then GoTo 29
                   If UCase$(FOREP1$) = "@HOY" Then GoTo 29
                   If InStr(FOREP1$, "{" + LOCKFI1$ + "}") > 0 Then GoTo 29 ' PARA QUE LOS CAMPOS ECO DEL PASSWORD PASEN
                   '
                   ' VLDX% = VALIVATE%(KI%)
                   VLDX% = VALIVATE%(KI%)
                   If VLDX% < 1 Then
                       On Error Resume Next
                       Text99(Index).BackColor = RGB(255, 255, 255)
                       On Error GoTo 0
                       Exit Sub
                    End If
                 End If
              End If
            End If
29        Next KI%
          '
Unload ENTRAPASS17
          ENTRAPASS17.Text2 = ""
' POSIBLE ERROR 14
          ENTRAPASS17.Left = Text99(Index).Left + Picture1.Left + Me.Left
          ENTRAPASS17.Top = Me.Top + Picture1.Top + Text99(Index).Top - ENTRAPASS17.Height - 100
          If ENTRAPASS17.Top < 1000 Then
             ENTRAPASS17.Top = Me.Top + Picture2.Top + Picture1.Top + Text99(Index).Top + 1000
          End If
          BANDENTRAPASS% = 1
          '
          For KKII% = 1 To Text99.Count
             Text99(KKII% - 1).BackColor = RGB(255, 255, 255)
          Next KKII%
          '
          Text99(Index).BackColor = RGB(100, 255, 100)
          Text99(Index) = ""
          
            ' MUEVE EL FORMULARIO
112         If Text99(Index).Top + Text99(Index).Height + 500 > Picture2.Height - Picture1.Top Then
               If VScroll1.Value < VScroll1.Max - 100 Then
                 VScroll1.Value = VScroll1.Value + 100
                 Call VScroll1_Scroll
                 GoTo 112
               End If
            End If
            '
113         If Text99(Index).Top < (-1) * Picture1.Top + 500 Then
               If VScroll1.Value >= 100 Then
                 VScroll1.Value = VScroll1.Value - 100
                 Call VScroll1_Scroll
                 GoTo 113
               End If
            End If
            '
          On Error Resume Next
          ENTRAPASS17.Show 1
          ENTRAPASS17.Visible = False
          BANDENTRAPASS% = 0
          If Err Then
             On Error GoTo 0
             Exit Sub
          End If
          On Error GoTo 0
          BANDENTRAPASS% = 0
          Text99(Index).BackColor = RGB(255, 255, 255)
          Text99(Index).TEXT = ""
          '
          TANTE = Timer
          VDEVUX$ = TRIM$(UCase$(ENTRAPASS17.Text1))
          If VDEVUX$ = "" Then
             On Error Resume Next
             Text99(Index) = ""
             Text99(Index).SetFocus
             On Error GoTo 0
             Exit Sub
          End If
          '
          If VDEVUX$ <> "" Then
             '
             ' VALIDACION DE LA CATEGORIA
             ID1$ = VDEVUX$
             CCAT$ = CCAT_OPER$(ID1$)
             CATI$ = CCAT$
             VALIMCAT = 0
             RESTOPA$ = TRIM$(Mid$(FOREP$, 10))
             If Left$(RESTOPA$, 1) = ">" Then
                   VALIMCAT = XVALO(Mid$(RESTOPA$, 2))
                   If VALIMCAT > 0 Then
                      If XVALO(CCAT$) < VALIMCAT Then
                         If CATI$ <> "" Then
                           Call MENSERR(24, "Categoría Informada del Operador (" & CATI$ & ") Menor a la Requerida (" & VALIMCAT & ").")
                           Exit Sub
                         End If
                      End If
                   End If
                 ElseIf Left$(RESTOPA$, 1) = "=" Then
                   RESTOPA$ = TRIM$(Mid$(RESTOPA$, 2))
                   RESTOPA$ = REPLACAR$(RESTOPA$, ",", ";")
                   RESTOPA$ = REPLACAR$(RESTOPA$, "/", ";")
                   RESTOPA$ = ";" + RESTOPA$
                   If InStr(RESTOPA$ + ";", ";" + CATI$ + ";") < 1 Then
                      Call MENSERR(24, "Categoría Informada del Operador (" & CATI$ & ") No Admisible Para Firma en esta Posición (" & Mid$(RESTOPA$, 2) & ").")
                      Exit Sub
                   End If
             End If
             ' FIN DE VALIDACION DE CATEGORIA
             '
             NOPXX$ = TRIM$(NOME_OPER$(VDEVUX$))
             If NOPXX$ = "" Then
                 TTXX$ = "El Identificador de Usuario '" & VDEVUX$ + "'"
                 TTXX$ = TTXX$ + "\No Está Asociado a Ningún Usuario del Sistema"
                 TTXX$ = TTXX$ + "\o Empleado del Padrón."
                 TTXX$ = TTXX$ + "\   \Realice la Asignación Correspondiente."
                 TTXX$ = TTXX$ + "\y Repita la Transaccion."
                 Call MENSERR(24, TTXX$)
             End If
             '
             If TRIM$(NOPXX$) <> "" Then   ' QUE SOLO BLOQUEE SI ENCUENTRA EL EMPLEADO
                TEXTO(ORITE%) = VDEVUX$                 ' ESTE TIENE QUE IR ANTES DEL VALIRANGO
                If VALIRANGO%(ORITE%) < 1 Then
                    Call MENSERR(24, "Usuario / Clave no Válido\para la Firma Electrónica Requerida.")
                    Exit Sub
                End If
                '
' POSIBLE ERROR 15
                Text99(Index).TEXT = VDEVUX$            ' ESTE VA DESPUES PARA NO ACTUALIZAR NOMBRE DE USUARIO SI NO VALIDA
                On Error Resume Next
'                If TRIM$(VDEVUX$) <> "" Then Text99(Index).Enabled = False
                On Error GoTo 0
                Text99(Index).BackColor = RGB(255, 255, 255)
                '
                ' aqui bloquea los campos asociados
                For KI% = 1 To Text99.Count
                  ORITE1% = ORITEXT99%(KI%)
                  If ORITE1% > 0 And ORITE1% <= CANOBJ% Then
                    LOCKFI1$ = UCase$(TRIM$(LOCKFI$(ORITE1%)))
                    If LOCKFI1$ <> "" Then
                      FOREP1$ = UCase$(FOREPLA$(ORITE1%))
                      If LOCKFI1$ = ACNAM$ Or InStr(FOREP1$, "{" + ACNAM$ + "}") > 0 Then
                        If Left$(UCase$(FOREP1$), 9) = "@PASSWORD" Then TREP = VDEVUX$: GoTo 90
                        If UCase$(FOREP1$) = "@USERNAME" Then TREP = USERNAME$: GoTo 90
                        If UCase$(FOREP1$) = "@AHORA" Then TREP = Time: GoTo 90
                        If UCase$(FOREP1$) = "@HOY" Then TREP = FECHATEX$(HOY(HOS$)): GoTo 90
                        '
                        If KI% <= Text99.Count Then
                          If TRIM$(Text99(KI%).TEXT) = "" Or Left$(FOREP1$, 7) = "APEYNOM" Or Left$(FOREP1$, 6) = "NOMBRE" Or Left$(FOREP1$, 6) = "APENOM" Or Left$(FOREP1$, 6) = "APYNOM" Then
                            If InStr(FOREP1$, "/") < 1 Then
                              TREP = TEXREPLA$(FOREP1$)
90                            TREPP = FORMATNUM(TREP, FORPA$(ORITE1%))
                              If TREPP <> TMASCARA(FORPA$(ORITE1%)) Then   ' si no se ha ingresado valor, mantiene vacio
                                '                                         '         y habilitado para ingresar
                                If TRIM$(CStr(TREPP)) <> "" Then                      ' para que no pise los ingresos manuales
                                  TEXTO(ORITE1%) = TREPP
                                  ' POSIBLE ERROR 16 - CAPTURADO
                                  On Error Resume Next
                                  Text99(KI%) = TREPP
                                  On Error GoTo 0
                                End If
                              End If
                            End If
                          End If
                          On Error Resume Next
                          If TRIM$(Text99(KI%).TEXT) <> "" Then Text99(KI%).Enabled = False
                          On Error GoTo 0
                        End If
                      End If
                    End If
                  End If
                  On Error Resume Next
                  Text99(KI% - 1).BackColor = RGB(255, 255, 255)
                  On Error GoTo 0
                Next KI%
             End If
             '
          End If
'        End If
' POSIBLE ERROR 17
        Text99(Index).BackColor = RGB(255, 255, 255)
      End If
      '
99    If LOCKFI$(ORITE%) <> "" Then
        If FOREP$ <> "" Then
          Call SETNEXTEX(ORITE%)
        End If
      End If
      '
    End If
    '
End Sub

Private Sub Text99_LostFocus_VIE(Index As Integer)
    COLOAN = Text99(Index).BackColor
    If ENTRAPASS17.Visible = False Then Text99(Index).BackColor = RGB(255, 255, 255)
    ORITE% = ORITEXT99%(Index)
    If ORITE% > 0 Then
      If ORITE% <= CANOBJ% Then
        If BANDENTRAPASS% = 1 Then
           Text99(Index).BackColor = RGB(100, 255, 100)
        End If
        '
        If TRIM$(Text99(Index)) = "" Then
           FORIXX$ = TRIM$(Left$(UCase$(FORPA$(ORITE%)), 1))
           If FORIXX$ = "F" Or FORIXX$ = "G" Or FORIXX$ = "$" Or FORIXX$ = "I" Then
                VALOR = 0
              Else
                VALOR = ""
           End If
           GoTo 25
        End If
        '
        If TRIM$(Text99(Index)) <> "" Then
          FORI$ = FORPA$(ORITE%)
          If Left$(FORI$, 1) = "F" Or Left$(FORI$, 1) = "G" Or Left$(FORI$, 1) = "$" Or Left$(FORI$, 1) = "I" Then
              If IsNumeric(Text99(Index)) = False Then
                Call MENSERR(24, "Ingreso No Numérico")
                On Error Resume Next
                Text99(Index).SetFocus
                On Error GoTo 0
                Exit Sub
              End If
              VALOR = XVALO(Text99(Index))
              '
            ElseIf Left$(FORI$, 1) = "D" Then
              FEX = FECHANUM(TRIM$(Text99(Index)))
              If FEX < 1 Then
                Call MENSERR(24, "Corresponde Ingreso Fecha")
                On Error Resume Next
                Text99(Index).SetFocus
                On Error GoTo 0
                Exit Sub
              End If
              VALOR = FEX
              '
            ElseIf Left$(FORI$, 1) = "H" Or Left$(FORI$, 1) = "T" Then
              VVXX$ = REPLACAR$(Text99(Index), ".", ":")
              VVXX$ = REPLACAR$(VVXX$, ",", ":")
              VVXX$ = REPLACAR$(VVXX$, "/", ":")
              VVXX$ = REPLACAR$(VVXX$, ";", ":")
              VALOR = VVXX$
            Else
              VALOR = TRIM$(Text99(Index))
          End If
          '
25        TEXTO(ORITE%) = VALOR
          If Left$(UCase$(FOREPLA$(ORITE%)), 9) = "@PASSWORD" Then
             GoTo 19 ' LOS PASSWORD SE VALIDAN AL GOTFOCUS, NO AL LOSTFOCUS
          End If
          '
          If VALIRANGO%(ORITE%) < 1 Then
              On Error Resume Next
              Text99(Index).BackColor = RGB(255, 100, 100)
              On Error GoTo 0
              If Left$(UCase$(FOREPLA$(ORITE%)), 9) = "@PASSWORD" Then
                   Call MENSERR(24, "Usuario / Clave no Válido\para la Firma Electrónica Requerida.")
                   On Error Resume Next
                   Text99(Index).SetFocus
                   On Error GoTo 0
                 Else
                   ' Call MENSERR(24, "Ingreso no Válido")
                   ' Call SETNEXTEX(ORITE%)
              End If
              Exit Sub
          End If
          '
19        CARMAN%(ORITE%) = 1
          Text99(Index) = FORMATNUM$(VALOR, FORI$)
          Call RECALCUTO(ORITE%)
        End If
      End If
      '
      On Error Resume Next
      If Left$(UCase$(FOREPLA$(ORITE%)), 9) <> "@PASSWORD" Then Text99(Index).BackColor = RGB(255, 255, 255)
      On Error GoTo 0
      '
    End If
    '
End Sub

Sub SETNEXTEX_VIE(NOBJ%)
    '
    INIBU% = NOBJ%
    '
    For KKII% = 1 To Text99.Count
       On Error Resume Next
       Text99(KKII% - 1).BackColor = RGB(255, 255, 255)
       On Error GoTo 0
    Next KKII%
    '
    If NOBJ% > 0 Then
       If FORMALTER.Visible = False Then
          MsgBox "Pulse 'Aceptar' para Habilitar Ingreso de Dato / Firma Electronica"
       End If
    End If
    '
15  PROTEX% = INIBU%
17  PROTEX% = PROTEX% + 1
    If PROTEX% > CANOBJ% Then
       If INIBU% > 0 Then
           INIBU% = 0
           GoTo 15
         Else
           Exit Sub
       End If
    End If
    If TIOBJ$(PROTEX%) <> "TEXT" Then GoTo 17
    If INVIS(PROTEX%) = 1 Then GoTo 17
    If LOCKFI$(PROTEX%) <> "" Then
       If FOREPLA$(PROTEX%) <> "" Then
          If UCase$(FOREPLA$(PROTEX%)) <> "@PASSWORD" Then
             GoTo 17
          End If
       End If
    End If
    '
    On Error GoTo 30
    If IOBJETX%(PROTEX%) > 0 Then
         Text99(IOBJETX%(PROTEX%)).SetFocus
         Call Text99_GotFocus(IOBJETX%(PROTEX%))
       ElseIf IOBJETXM%(PROTEX%) > 0 Then
         Text99M(IOBJETXM%(PROTEX%)).SetFocus
         Call Text99M_GotFocus(IOBJETX%(PROTEX%))
    End If
26  On Error GoTo 0
    Exit Sub
    '
30  Resume 31
31  On Error GoTo 0
    GoTo 17
    '
End Sub

Private Sub Timer1_Timer()

   NUCETE$ = TRIM$(Label6)
   If NUCETE$ = "" Then
     Exit Sub
   End If
   '
   LETR$ = "N"
   On Error Resume Next
   AppActivate "Save PDF"
   If Err Then
     LETR$ = "O"
     On Error GoTo 0
     On Error Resume Next
     AppActivate "Guardar como"
     If Err Then
        On Error GoTo 0
        Exit Sub
     End If
   End If
   On Error GoTo 0
   '
   'Se modifica por que en la ventana de windows seven de guardar como la letra es 'O' en vez de N
   If LETR$ = "N" Then SendKeys "%N" + LUCOM$ + "" + NUCETE$ + ".PDF" + Chr$(13)

   If LETR$ = "O" Then SendKeys "%o" + LUCOM$ + "" + NUCETE$ + ".PDF" + Chr$(13)

   NUCETE$ = ""
   Timer1.Enabled = False
   Exit Sub
   '
End Sub
