VERSION 5.00
Begin VB.Form INITAB04 
   BackColor       =   &H00D0F0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reportes Tabulados"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   15330
   ClientWidth     =   5685
   Icon            =   "INITAB04.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   5685
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404080&
      Height          =   4335
      Left            =   4800
      ScaleHeight     =   4275
      ScaleWidth      =   915
      TabIndex        =   0
      Top             =   0
      Width           =   975
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   100
         ScaleHeight     =   75
         ScaleWidth      =   540
         TabIndex        =   3
         Top             =   3000
         Width           =   600
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
      Begin VB.CommandButton Command29 
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
         Left            =   100
         Picture         =   "INITAB04.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   850
         Width           =   650
      End
      Begin VB.CommandButton Command24 
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
         Height          =   640
         Left            =   100
         Picture         =   "INITAB04.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   210
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -320
         Picture         =   "INITAB04.frx":075E
         Top             =   3320
         Width           =   1515
      End
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5880
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "INITAB04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command24_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Call FINAL("")
End Sub

Private Sub Form_Load()
   '
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")
   '
   Top = Screen.Height + 1000
   '
   Call RECLAPAG   ' RECLAMO DE PAGO - se deshabilita - va a acreedores y ordenes de compra
   '
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
      Caption = Caption + "  -  " + EPAC$
   End If
   Call GRATITFOR(Caption)
   '
   Call SELPROCE
   '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Call FINAL("")
End Sub

Sub SELPROCE()
   '
   COMMA$ = APLINVO$
   PPX% = InStr(COMMA$, "/"): If PPX% < 1 Then PPX% = 1 + Len(COMMA$)
   PROPRIN$ = TRIM$(Left$(COMMA$, PPX% - 1))
   ATRIBU$ = TRIM$(Mid$(COMMA$, PPX% + 1))
   '
'PROPRIN$ = "*LIPRESU"
'MsgBox PROPRIN$
'ATRA$ = CurDir
   If Left$(PROPRIN$, 1) = "*" Then
     PROPRIN$ = Mid$(PROPRIN$, 2)
   End If
   '
   If PROPRIN$ = "" Then
     Call CENTRAFORM(Me)
     Show
     Exit Sub
   End If
   '
   Show
   If ATRA$ = "" Then ATRA$ = App.Path
   '
   FORMA$ = ATRA$ + "\" + PROPRIN$ + ".PRF"
   
   If EXISTE%(FORMA$) < 1 Then
     Call EXTRACT(FORMA$)
   End If
   '
   If EXISTE%(FORMA$) < 1 Then
     FORMB$ = REPLACAR$(FORMA$, "\", "/")
     Call COMUNI("Proceso Interrumpido.\  \No se ha Podido Hallar o Extraer\el Controlador Correspondiente\al Proceso Elegido:\  \" + FORMB$ + "\ \Consulte a su Soporte de Sistemas.")
     Call FINAL("")
   End If
   '
   APROP$ = LOADFILE$(FORMA$)
   XX$ = Left$(APROP$, 128)      ' PRIMER REGISTRO
   TIPRO% = CVI(Mid$(XX$, 127, 2))
   '
   Screen.MousePointer = 1
   '
   Select Case TIPRO%
     '
     Case 1              ' SELECCIONA EDICION TABULADA
        Call CARDATFORM
        Call DIBUFORM
        YAFIL% = 0
        WEDTAB04.Show 1
        If PROPRIN$ <> "" Then Call FINAL("")
        Exit Sub
     Case 2              ' SELECCIONA LISTADO TABULADO
        WLITAB04.Show 1
        If PROPRIN$ <> "" Then Call FINAL("")
        Exit Sub
     Case 3              ' SELECCIONA FILTRO TABULADO
           'CCI$(0) = PROPRIN$
           'Call CONECRUN("#FILTABU\" + PROPRIN$, "")
           'Exit Sub
 '              WFILTAB.Show
 '              WFILTAB.Refresh
               Exit Sub
     Case 4              ' SELECCIONA ORDENAMIENTO
        Call CARDATSOR
        WSORTB04.Show 1
        If PROPRIN$ <> "" Then Call FINAL("")
        Exit Sub
     Case 5              ' SELECCIONA PANTALLA PLANA
 '              WPLATAB.Show
 '              WPLATAB.Refresh
            Exit Sub
   End Select
   '
   Call MENSERR(24, "No se Encuentra Controlador\Correspondiente a " + PROPRIN$ + ".PRF")
   Call FINAL("")
   '
End Sub
