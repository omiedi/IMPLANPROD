VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form PRINMAI09 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Muestra Mensaje de Correo Electrónico Recibido"
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9375
   BeginProperty Font 
      Name            =   "Small Fonts"
      Size            =   6
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   9375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Caption         =   "Solicita "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   2550
      TabIndex        =   26
      Top             =   7560
      Width           =   1280
      Begin VB.CheckBox Check1 
         Caption         =   "Confirm"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   27
         Top             =   350
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Acciones Correctivas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2650
      Left            =   3960
      TabIndex        =   21
      Top             =   5670
      Width           =   5340
      Begin VB.Frame Frame8 
         Caption         =   "Borrado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   130
         TabIndex        =   33
         Top             =   1575
         Width           =   3540
         Begin VB.OptionButton Option5 
            Caption         =   "Borrado Directo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   420
            TabIndex        =   35
            Top             =   315
            Width           =   3000
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Actualizar Reglas de Borrado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   420
            TabIndex        =   34
            Top             =   550
            Width           =   3000
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Registrar Remitente"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Left            =   130
         TabIndex        =   28
         Top             =   315
         Width           =   5055
         Begin VB.CheckBox Check3 
            Caption         =   "Guardar Adjuntos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2730
            TabIndex        =   32
            Top             =   840
            Value           =   1  'Checked
            Width           =   2025
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Guardar E-Mail"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   420
            TabIndex        =   31
            Top             =   840
            Value           =   1  'Checked
            Width           =   2025
         End
         Begin VB.OptionButton Option1 
            Caption         =   "En Ficha de Contacto Existente"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   735
            TabIndex        =   30
            Top             =   315
            Value           =   -1  'True
            Width           =   4000
         End
         Begin VB.OptionButton Option2 
            Caption         =   "En Nueva Ficha de Contacto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   735
            TabIndex        =   29
            Top             =   525
            Width           =   4000
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Continuar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   3850
         TabIndex        =   23
         Top             =   1640
         Width           =   1320
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Actualizar Reglas de Borrado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   420
         TabIndex        =   22
         Top             =   1575
         Width           =   4800
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   24
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   8610
      ScaleHeight     =   75
      ScaleWidth      =   585
      TabIndex        =   14
      Top             =   1100
      Width           =   650
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Recibido en "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   105
      TabIndex        =   12
      Top             =   7560
      Width           =   2385
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   350
         Width           =   1935
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Archivos Adjuntos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1880
      Left            =   105
      TabIndex        =   6
      Top             =   5670
      Width           =   3735
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1420
         IntegralHeight  =   0   'False
         Left            =   105
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   315
         Width           =   3510
      End
   End
   Begin VB.CommandButton Command7 
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
      Height          =   675
      Left            =   8600
      Picture         =   "PRINMAI09.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Cierra y Abandona la Aplicación"
      Top             =   210
      Width           =   645
   End
   Begin VB.Frame Frame1 
      Caption         =   "Origen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   105
      Width           =   8350
      Begin VB.Label Label17 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   19
         Top             =   960
         Width           =   4065
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Id:"
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
         Left            =   5040
         TabIndex        =   37
         Top             =   1005
         Width           =   555
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5670
         TabIndex        =   36
         Top             =   945
         Width           =   2530
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7440
         TabIndex        =   25
         Top             =   1260
         Width           =   765
      End
      Begin VB.Label Label16 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   18
         ToolTipText     =   "Doble-Click para Incluir en Regla de Borrado"
         Top             =   630
         Width           =   7005
      End
      Begin VB.Label Label15 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   17
         ToolTipText     =   "Doble-Click para Incluir en Regla de Borrado"
         Top             =   1275
         Width           =   6165
      End
      Begin VB.Label Label14 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   16
         ToolTipText     =   "Doble-Click para Incluir en Regla de Borrado"
         Top             =   300
         Width           =   5205
      End
      Begin VB.Label Label26 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6960
         TabIndex        =   11
         Top             =   300
         Width           =   1245
      End
      Begin VB.Label Label43 
         Alignment       =   1  'Right Justify
         Caption         =   "Cte:"
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
         Left            =   6360
         TabIndex        =   10
         Top             =   360
         Width           =   555
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Para:"
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
         Left            =   320
         TabIndex        =   9
         Top             =   1030
         Width           =   750
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         Caption         =   "Asunto:"
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
         TabIndex        =   4
         Top             =   1365
         Width           =   960
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "De:"
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
         Left            =   210
         TabIndex        =   3
         Top             =   405
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Dir-elec:"
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
         Left            =   315
         TabIndex        =   2
         Top             =   720
         Width           =   750
      End
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   735
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   327680
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   327680
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.Frame Frame2 
      Caption         =   "Detalle del Mensaje Recibido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3825
      Left            =   105
      TabIndex        =   0
      Top             =   1850
      Width           =   9180
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   3450
         Left            =   105
         TabIndex        =   39
         Top             =   315
         Visible         =   0   'False
         Width           =   8895
         ExtentX         =   15690
         ExtentY         =   6085
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   735
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   38
         Text            =   "PRINMAI09.frx":014A
         Top             =   525
         Visible         =   0   'False
         Width           =   8910
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Text            =   "PRINMAI09.frx":0152
         Top             =   315
         Width           =   8910
      End
   End
   Begin VB.Image Image4 
      Height          =   390
      Left            =   8175
      Picture         =   "PRINMAI09.frx":0158
      Top             =   1380
      Width           =   1515
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   200
      TabIndex        =   5
      Top             =   1890
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "PRINMAI09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RESTODELMAIL$
Dim ENVIOMASIVO%

Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
  
Private Sub Command16_Click()
    Call SELECHO("SELFECHA")
    VDEV$ = VALACT1$("SELFECHA")
    If VDEV$ <> "" Then
        Text7.TEXT = VDEV$
    End If
End Sub

Private Sub Command1_Click()
   Label1 = ""
   CONDI$ = "NUME_MAIL = " & TRIM$(Label13)
   If Option1.Value = True Then
        Call CARGAENFICHAEXISTENTE
        DIELREM$ = TRIM$(Label16)
        Label26 = CODPROSPE$(DIELREM$)
        If TRIM$(Label26) <> "" Then Hide
      ElseIf Option2.Value = True Then
        Call NUEVAFICHADECONTACTO
        DIELREM$ = TRIM$(Label16)
        Label26 = CODPROSPE$(DIELREM$)
        If TRIM$(Label26) <> "" Then Hide
      ElseIf Option4.Value = True Then
        ABUX1$ = TRIM$(Label14) + "<" + TRIM$(Label16) + ">"
        ABUX2$ = TRIM$(Label15)
        CRITBORR09.Text1(0) = ABUX1$
        CRITBORR09.Combo1(0) = "Remitente"
        CRITBORR09.Text1(1) = ABUX2$
        CRITBORR09.Combo1(1) = "Asunto"
        CRITBORR09.Show 1
        Label1 = 2
        Hide
      ElseIf Option5.Value = True Then
        Label26 = "DELETE"
        Label1 = 1
        Hide
   End If
   '
End Sub

Sub NUEVAFICHADECONTACTO()
   '
   ACTU% = 0
   FISECOM09.Label26 = ""
   Call FISECOM09.Command21_Click
   If FISECOM09.Label26 <> "" Then
      FISECOM09.Text5 = Label16
      FISECOM09.Text6(0) = Label14
      FISECOM09.Text8(0) = Label16
      ACTU% = 1
      FISECOM09.Show 1
   End If
   '
   Label1 = ACTU%
   '
End Sub

Sub CARGAENFICHAEXISTENTE()
   '
   ABUSX$ = "INDIPROS" ' : If USNBR% < 100 Then ABUSX$ = "BSEGALE"
21 Call SELECHO(ABUSX$ & "/Indice Prospect")
   ACTU% = 0
   VDEVU$ = VALACT1$(ABUSX$)
   If VDEVU$ <> "" Then
      FISECOM09.Label26 = TRIM$(VDEVU$)
      NOREMI$ = TRIM$(Label14)
      DRELRE$ = TRIM$(Label16)
      If FISECOM09.Text5 = "" Then
         FISECOM09.Text5 = DRELRE$
         ACTU% = 1
      End If
35    For KKI% = 1 To FISECOM09.Text6.Count
         If TRIM$(FISECOM09.Text6(KKI% - 1)) = "" Then
           If TRIM$(FISECOM09.Text8(KKI% - 1)) = "" Then
             FISECOM09.Text6(KKI% - 1) = NOREMI$
             FISECOM09.Text8(KKI% - 1) = DRELRE$
             ACTU% = 1
             GoTo 80
           End If
         End If
         If UCase$(TRIM$(FISECOM09.Text6(KKI% - 1))) = UCase$(NOREMI$) Then
           If UCase$(TRIM$(FISECOM09.Text8(KKI% - 1))) = UCase$(DRELRE$) Or TRIM$(FISECOM09.Text8(KKI% - 1)) = "" Then
             FISECOM09.Text8(KKI% - 1) = LCase$(DRELRE$)
             ACTU% = 1
             GoTo 80
           End If
         End If
      Next KKI%
      '
      Call FISECOM09.AGREGARENGLONES
      GoTo 35
      '
80    If ACTU% > 0 Then
         Call FISECOM09.Command1_Click
         FISECOM09.Show 1
      End If
      '
   End If
   Label1 = ACTU%
   '
End Sub
'
Private Sub Command7_Click()
   'Call CIERRARCH("*.*")
   Label1 = "0"
   Hide
End Sub

Private Sub Command9_Click()
   '
   Call SELECHO("SELFECHA")
   VDEV$ = VALACT1$("SELFECHA")
   If VDEV$ <> "" Then
       Label4.Caption = VDEV$
   End If
   '
End Sub

Private Sub Form_Activate()
   If ATENTI% > 0 Then
     Left = 0
     Top = 0
     'Call CENTRAFORM(Me)
     ATENTI% = 0
     Text2.SetFocus
   End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UTIMETEC = Timer
End Sub

Private Sub Form_Unload(Cancel As Integer)
'   Call CIERRARCH("*.*")
'   Call BAJALDISCO
   Cancel = 0
End Sub

Private Sub Label13_Change()
    WebBrowser1.Visible = False
    NUMAI& = XVALO(Label13)
    CONDI$ = "NUME_MAIL = " & NUMAI&
    If CantRegistros("BANPENPRO", CONDI$, "NOMESS") > 0 Then
          '
          Caption = "Muestra Mensaje de Correo Recibido#" & TRIM$(Str$(NUMAI&))
          Label26 = TRIM$(VALOBADA("BANPENPRO", "PROS_DEST", CONDI$, "NOMESS"))
          Label17 = TRIM$(VALOBADA("BANPENPRO", "PARA_MAIL", CONDI$, "NOMESS"))
          Label4 = TRIM$(VALOBADA("BANPENPRO", "IDEN_MAIL", CONDI$, "NOMESS"))
          Label15 = TRIM$(VALOBADA("BANPENPRO", "ASUN_MAIL", CONDI$, "NOMESS"))
          Label16 = TRIM$(VALOBADA("BANPENPRO", "DREL_MAIL", CONDI$, "NOMESS"))
          Label14 = TRIM$(VALOBADA("BANPENPRO", "ORIG_MAIL", CONDI$, "NOMESS"))
          '
          TXMAIL$ = VALOBADA("BANPENPRO", "TEXT_MAIL", CONDI$, "NOMESS")
          HTMAIL$ = VALOBADA("BANPENPRO", "HTMT_MAIL", CONDI$, "NOMESS")
          '
          If TRIM$(HTMAIL$) <> "" Then
               Call SAVEFILE("C:\FLEXOFT\ARADJUN\HTMMAIL.HTM", HTMAIL$)
               WebBrowser1.Navigate "C:\FLEXOFT\ARADJUN\HTMMAIL.HTM"
               WebBrowser1.Visible = True
            Else
               Text1.TEXT = TXMAIL$
               Text1.Refresh
          End If
          '
          Check1.Value = XVALO(VALOBADA("BANDENTRA", "SOLI_CONF", CONDI$, "NOMESS"))
          Call CARGALISTA(List3, "ADJPENPRO", "NORD_MAIL/F3; FILE_NAME/A128", CONDI$, "NOMESS")
          '
       Else
          Call MENSERR(24, "Imposible Abrir Mensaje Enviado Numero '" & TRIM$(Str$(NUMAI&)) & "'")
   End If
'    Call ACTUREGISTRO("BANPENPRO", "FHOR_RECEP", CONDI$, FRECEP, RAFE&, "NOMESS")
End Sub

Private Sub Label4_DblClick()
   Label4 = ""
   Text1 = ""
End Sub

Private Sub Label14_DblClick()
    CRITBORR09.Text1(0) = TRIM$(Label14)
    CRITBORR09.Combo1(0) = "Remitente"
    CRITBORR09.Show 1
    Label1 = "2"
    Hide
End Sub

Private Sub Label15_DblClick()
    CRITBORR09.Text1(0) = TRIM$(Label15)
    CRITBORR09.Combo1(0) = "Asunto"
    CRITBORR09.Show 1
    Label1 = "2"
    Hide
End Sub

Private Sub Label16_DblClick()
    CRITBORR09.Text1(0) = TRIM$(Label16)
    CRITBORR09.Combo1(0) = "Remitente"
    CRITBORR09.Show 1
    Label1 = "2"
    Hide
End Sub

Private Sub List3_DblClick()
   On Error Resume Next
   MkDir LUCOM$ + "ARADJUN"
   On Error GoTo 0
   RUTAR$ = LUCOM$ + "ARADJUN\"
   '
   NUMAI& = XVALO(Label13)
   FINAM$ = TRIM$(Mid$(List3.TEXT, 129))
   XXXX$ = LOADFILE$(FINAM$)
   While InStr(FINAM$, "\") > 0: FINAM$ = Mid$(FINAM$, 2): Wend
   '
   If FINAM$ <> "" Then
      If XXXX$ <> "" Then
         Call SAVEFILE(RUTAR$ + FINAM$, XXXX$)
         DOCUACT$ = RUTAR$ + FINAM$
         Call MUESTRADOC4(DOCUACT$)
      End If
   End If
   '
End Sub

Sub MUESTRADOC4(DOCUACT$)
    '
    Call ShellExecute(Me.hwnd, "Open", DOCUACT$, "", "", 1)
    '
End Sub

Private Sub Text12_DblClick()
    Call SELECHO("SELFECHA")
    VDEV$ = VALACT1$("SELFECHA")
    If VDEV$ <> "" Then
        Text12.TEXT = VDEV$
    End If
End Sub

Sub ARMADESTI(CODICLI$, Optional DESTI$)
   '
   KOKLI$ = TRIM$(CODICLI$)
   CONDICION$ = "CODI_CLI = '" & KOKLI$ & "'"
   '
   EMPRE$ = TRIM$(VALOBADA("BSEGCOM", "RASO_CLI", CONDICION$)) 'TRIM$(Mid$(RELE$, 3, 59))
   CALNUM$ = TRIM$(VALOBADA("BSEGCOM", "DOMI_CLI", CONDICION$)) 'TRIM$(Mid$(RELE$, 129, 48))
   CPOST$ = TRIM$(VALOBADA("BSEGCOM", "Cpos_Cli", CONDICION$)) 'TRIM$(Mid$(RELE$, 177, 8))
   If CPOST$ <> "" Then CPOST$ = CPOST$ + " "
   LOCAX$ = CPOST$ + TRIM$(VALOBADA("BSEGCOM", "Loca_Cli", CONDICION$)) ' CPOST$ + TRIM$(Mid$(RELE$, 185, 32))
   PROPA$ = TRIM$(VALOBADA("BSEGCOM", "Prov_Cli", CONDICION$)) 'TRIM$(Mid$(RELE$, 217, 32))
   PRO$ = TRIM$(ECOARCH$("provinc", PROPA$))
   If PRO$ <> "" Then PRO$ = PRO$ + " - "
   
   PRO$ = PRO$ + TRIM$(VALOBADA("BSEGCOM", "Pais_Cli", CONDICION$)) 'TRIM$(PROPA$ + TRIM$(Mid$(RELE$, 249, 24)))
   '
   TTXX$ = ""
   If EMPRE$ <> "" Then TTXX$ = TTXX$ + EMPRE$ + Chr$(13) + Chr$(10)
   If CALNUM$ <> "" Then TTXX$ = TTXX$ + CALNUM$ + Chr$(13) + Chr$(10)
   If LOCAX$ <> "" Then TTXX$ = TTXX$ + LOCAX$ + Chr$(13) + Chr$(10)
   If PRO$ <> "" Then TTXX$ = TTXX$ + PRO$ + Chr$(13) + Chr$(10)
   '
   ASUNTX$ = TRIM$(ASUNTO$)
   If ASUNTX$ = "" Then ASUNTX$ = CONTROL$("SEGUICOM", "ASUNTO")
   If ASUNTX$ = "" Then ASUNTX$ = "FLEXOFT"
   If ASUNTX$ = "FLEXOFT" Then
      If TRIM$(DESTI$) <> "" Then
         ASUNTX$ = ASUNTX$ + " para " + DESTI$
      End If
   End If
   '
   SHOWMAIL09.Label15 = ASUNTX$
   '
   HOII% = HOY(HOS$)
   FEC = FETEXTO(HOII%)
   LUGARYFECH$ = AJUSTD("Buenos Aires - " + FEC, 100)
   TTYY$ = LUGARYFECH$ + vbCrLf
   '
   TTYY$ = TTYY$ + "Sr./Sra." + Chr$(13) + Chr$(10)
   TTYY$ = TTYY$ + DESTI$ + Chr$(13) + Chr$(10)
   '
   TTYY$ = TTYY$ + TTXX$ + RESTODELMAIL$
   Text2 = TTYY$
   '
End Sub

Private Sub Option1_Click()
   If Option1.Value = True Then
      Option4.Value = False
      Option5.Value = False
      Check2.Value = 1
      Check3.Value = 1
   End If
End Sub

Private Sub Option1_DblClick()
   Call Command1_Click
End Sub

Private Sub Option2_Click()
   If Option2.Value = True Then
      Option4.Value = False
      Option5.Value = False
      Check2.Value = 1
      Check3.Value = 1
   End If
End Sub

Private Sub Option2_DblClick()
   Call Command1_Click
End Sub

Private Sub Option4_Click()
   If Option4.Value = True Then
      Option1.Value = False
      Option2.Value = False
      Check2.Value = 0
      Check3.Value = 0
   End If
End Sub

Private Sub Option4_DblClick()
   Call Command1_Click
End Sub

Private Sub Option5_Click()
   If Option5.Value = True Then
      Option1.Value = False
      Option2.Value = False
      Check2.Value = 0
      Check3.Value = 0
   End If
End Sub

Private Sub Option5_DblClick()
   Call Command1_Click
End Sub
