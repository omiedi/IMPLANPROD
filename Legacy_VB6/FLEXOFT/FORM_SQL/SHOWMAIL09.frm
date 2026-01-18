VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form SHOWMAIL09 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Muestra Mensaje de Correo Electrónico"
   ClientHeight    =   7890
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
   ScaleHeight     =   7890
   ScaleWidth      =   9375
   ShowInTaskbar   =   0   'False
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
      Height          =   570
      Left            =   8600
      Picture         =   "SHOWMAIL09.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "Re-Enviar a Otro Destinatario"
      Top             =   4080
      Width           =   645
   End
   Begin VB.Frame Frame5 
      Caption         =   "Preparó"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   3465
      TabIndex        =   40
      Top             =   5880
      Width           =   2250
      Begin VB.Label Label10 
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
         Left            =   170
         TabIndex        =   42
         Top             =   630
         Width           =   2000
      End
      Begin VB.Label Label20 
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
         Left            =   170
         TabIndex        =   41
         Top             =   280
         Width           =   2000
      End
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
      Height          =   570
      Left            =   8600
      Picture         =   "SHOWMAIL09.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Envío Multiple"
      Top             =   2760
      Width           =   645
   End
   Begin VB.CommandButton Command18 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   8600
      Picture         =   "SHOWMAIL09.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Anulacion de Mail Pendiente de Envio"
      Top             =   3360
      Width           =   645
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   8620
      ScaleHeight     =   75
      ScaleWidth      =   585
      TabIndex        =   30
      Top             =   5085
      Width           =   650
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   31
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Ultimo Envio Realizado"
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
      Left            =   5800
      TabIndex        =   26
      Top             =   6980
      Width           =   3500
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
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
         Left            =   2940
         TabIndex        =   29
         Top             =   315
         Width           =   390
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Veces:"
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
         Left            =   2100
         TabIndex        =   28
         Top             =   360
         Width           =   750
      End
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
         Left            =   195
         TabIndex        =   27
         Top             =   315
         Width           =   1935
      End
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   8600
      Picture         =   "SHOWMAIL09.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Re-Enviar Inmediatamente"
      Top             =   1215
      Width           =   645
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
      Left            =   120
      TabIndex        =   12
      Top             =   5880
      Width           =   3250
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
         TabIndex        =   13
         Top             =   315
         Width           =   3040
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
      Height          =   885
      Left            =   8600
      Picture         =   "SHOWMAIL09.frx":0D60
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Cierra y Abandona la Aplicación"
      Top             =   210
      Width           =   645
   End
   Begin VB.Frame Frame10 
      Caption         =   "Envio Programado"
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
      Left            =   3465
      TabIndex        =   8
      Top             =   6980
      Width           =   2265
      Begin VB.Label Label19 
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
         Left            =   1260
         TabIndex        =   43
         Top             =   315
         Width           =   675
      End
      Begin VB.Label Label4 
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
         Left            =   120
         TabIndex        =   10
         Top             =   315
         Width           =   990
      End
      Begin VB.Label Label1 
         Caption         =   "Hs"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1965
         TabIndex        =   9
         Top             =   285
         Width           =   225
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   5800
      TabIndex        =   7
      Top             =   5880
      Width           =   1815
      Begin VB.CommandButton COMMAND16 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Left            =   1560
         TabIndex        =   22
         ToolTipText     =   "Lista Completa de Contactos"
         Top             =   600
         Width           =   150
      End
      Begin VB.TextBox Text7 
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
         Height          =   250
         Left            =   600
         TabIndex        =   20
         Top             =   600
         Width           =   960
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   1080
         TabIndex        =   18
         Top             =   240
         Width           =   360
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Repetir c/"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Hasta:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   21
         Top             =   630
         Width           =   465
      End
      Begin VB.Label Label8 
         Caption         =   "ds"
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
         Left            =   1470
         TabIndex        =   19
         Top             =   270
         Width           =   225
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Confirmacion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   7700
      TabIndex        =   6
      Top             =   5880
      Width           =   1610
      Begin VB.CheckBox Check1 
         Caption         =   "Solicitada"
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
         TabIndex        =   16
         Top             =   280
         Width           =   1275
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   80
         TabIndex        =   37
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Destinatario"
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
      TabIndex        =   2
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
         TabIndex        =   36
         Top             =   960
         Width           =   7005
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
         TabIndex        =   35
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
         TabIndex        =   34
         Top             =   1270
         Width           =   7005
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
         TabIndex        =   33
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
         TabIndex        =   24
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
         TabIndex        =   23
         Top             =   360
         Width           =   555
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "C/Copia:"
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
         TabIndex        =   15
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
         TabIndex        =   5
         Top             =   1365
         Width           =   960
      End
      Begin VB.Label Label3 
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
         Left            =   210
         TabIndex        =   4
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
         TabIndex        =   3
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
      Caption         =   "Detalle de Mensaje Personalizado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3930
      Left            =   105
      TabIndex        =   0
      Top             =   1850
      Width           =   8340
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
         Height          =   3510
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Text            =   "SHOWMAIL09.frx":0EAA
         Top             =   315
         Width           =   8100
      End
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
      Left            =   8600
      TabIndex        =   32
      Top             =   2280
      Width           =   645
   End
   Begin VB.Image Image4 
      Height          =   390
      Left            =   8175
      Picture         =   "SHOWMAIL09.frx":0EB2
      Top             =   5355
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
      TabIndex        =   11
      Top             =   1890
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "SHOWMAIL09"
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
        
Sub Command1_Click()
   '
   ' crea los adjuntos
   On Error Resume Next
   MkDir LUCOM$ + "ARADJUN"
   On Error GoTo 0
   RUTAR$ = LUCOM$ + "ARADJUN\"
   On Error Resume Next
   Kill RUTAR$ + "\*.*"
   On Error GoTo 0
   '
   For i& = 1 To 256: MAI_ADJUN$(i&) = "": Next i&
   NUMAI& = XVALO(Label13)
   JJ& = 0
   SQLX$ = "SELECT * FROM ADJUNMAI WITH(NOLOCK) WHERE NUME_MAIL = " & NUMAI& & " ORDER BY NORD_MAIL ASC"
   Dim ADJUN As ADODB.Recordset
   Set ADJUN = New ADODB.Recordset
   Set ADJUN = FLEXCONN.Execute(FILTSQL$(SQLX$))
   With ADJUN
      Do While Not .EOF
         FINAM$ = REPLACAR$(SAFETEXT$(!FILE_NAME), " ", "_")
         XXXX$ = SAFETEXT$(!FILE_DATA)
         If FINAM$ <> "" Then
           If XXXX$ <> "" Then
             Call SAVEFILE(RUTAR$ + FINAM$, XXXX$)
             JJ& = JJ& + 1
             MAI_ADJUN$(JJ&) = RUTAR$ + FINAM$
           End If
         End If
      .MoveNext
      Loop
   End With
   '
   TTXX$ = LOADFILE(LUCOM$ + "STPFMAIL.CFG")
   QUEDAPEN% = 0: If TRIM$(Mid$(TTXX$, 553, 8)) = "CLIENT" Then QUEDAPEN% = 1
   Call CREACAMPO("", "BANDENVI", "PENDIENV", 3, 2, 1)
   NUMAI& = XVALO(Label13)
   CONDI$ = "NUME_MAIL = " & NUMAI&
   Call ACTUREGISTRO("BANDENVI", "PENDIENV", CONDI$, QUEDAPEN%, RAFE&, "NOMESS")
   '
   ' falta validar los datos
   Call ENVIAMAIL
   If ENVIOMASIVO% > 0 Then Exit Sub
   '
   '
   NUMAI& = XVALO(Label13)
   CONDI$ = "NUME_MAIL = " & NUMAI&
   HOII% = HOY(HO$)
   ' AHORA = FETEXCOR1$(HOII%) + " " + Time$
   '
   FINPROGRA% = FECHANUM(Text7)
   FREPI% = XVALO(Text6): If Check2.Value < 1 Then FREPI% = 0
   PROENVI% = FECHANUM("01/01/80")
   If FREPI% > 0 Then
      If DIASPOST(HOII%, FREPI%) <= FINPROGRA% Then
        PROENVI% = DIASPOST(HOII%, FREPI%)
      End If
   End If
   PROEN = FETEXCOR1$(PROENVI%)
   If PROENVI% >= HOII% Then PROEN = PROEN + " " + Time$
   '
   Call ACTUREGISTRO("BANDENVI", "FHOR_ULTENVI", CONDI$, AHORA, RAFE&, "NOMESS")
      NUMAI& = XVALO(Label13)
      CONDI$ = "NUME_MAIL = " & NUMAI&
      ENVIANT& = XVALO(VALOBADA("BANDENVI", "CANT_ENVI", CONDI$, "NOMESS"))
   Call ACTUREGISTRO("BANDENVI", "CANT_ENVI", CONDI$, 1 + ENVIANT&, RAFE&, "NOMESS")
   Call ACTUREGISTRO("BANDENVI", "FHOR_SOLENVI", CONDI$, PROEN, RAFE&, "NOMESS")    ' proximo envio proyectado
      '
      SOLICO% = Check1.Value
   Call ACTUREGISTRO("BANDENVI", "SOLI_CONF", CONDI$, SOLICO%, RAFE&, "NOMESS")
   Call ACTUREGISTRO("BANDENVI", "FREC_REPI", CONDI$, XVALO(Check2) * XVALO(Text6), RAFE&, "NOMESS")
   Call ACTUREGISTRO("BANDENVI", "FEFIN_REPI", CONDI$, FETEXCOR1$(FECHANUM(Text7)), RAFE&, "NOMESS")
   '
   '
   ' ANOTA EN HISTORIAL DE SEGUIMIENTOS
   If ENVIAINMEDIATO% < 1 Then
      COCLI$ = TRIM$(Label26)
      CODOP$ = TRIM$(Label7)
      If COCLI$ <> "" Then
         TTTT$ = "Envío de Mail (#" & TRIM$(Label13) & "#) " & Format(Now, "dd/mm/yy - hh:mm") + CRLF$ + CRLF$
         Call FISECOM09.ANOTAHISTORIAL(TTTT$, COCLI$)
      End If
   End If
   '
   Call CIERRARCH("*.*")
   On Error Resume Next
   Unload Me
   On Error GoTo 0
   '
End Sub

Sub ENVIAMAIL()
      
   MAI_DESTI$ = TRIM$(Label14)
   MAI_DIREL$ = TRIM$(Label16)
   MAI_COPIA$ = TRIM$(Label17)
   MAI_ASUNT$ = TRIM$(Label15)
   Call FRATEXTO("." & Text2, 88, 1, 1)
   TTXX$ = ""
   For KKI% = 1 To CARETEX%
     TTXX$ = TTXX$ + RETEX$(KKI%) + CRLF$
   Next KKI%
   TTYY$ = REPLACAR$(TTXX$, Chr$(13), "~")
   TTYY$ = REPLACAR$(TTXX$, Chr$(10), "")
   TTYY$ = REPLACAR$(TTXX$, "(", "-")
   TTYY$ = REPLACAR$(TTXX$, ")", "-")
   MAI_TEXTO$ = TTYY$
   '
   Call GENERAMAIL
   '
End Sub

Private Sub Command16_Click()
    Call SELECHO("SELFECHA")
    VDEV$ = VALACT1$("SELFECHA")
    If VDEV$ <> "" Then
        Text7.TEXT = VDEV$
    End If
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   Load SELDESTI09
   If SELDESTI09.List5.ListCount < 1 Then
      Call MENSERR(24, "Envio Masivo Imposible.\No hay Grupos Definidos.")
      GoTo 99
   End If
   '
   PPXX% = InStr(UCase$(Text2), "DE NUESTRA MAYOR CONSIDE")
   If PPXX% < 1 Then
      PPXX% = InStr(UCase$(Text2), "ESTIMADO")
      If PPXX% < 1 Then
         Call MENSERR(24, "Imposible Armar Texto del Mail")
         GoTo 99
      End If
   End If
   RESTODELMAIL$ = Mid$(Text2, PPXX%)
   '
   SELDESTI09.Show 1
   OPENFORMS = DoEvents
   If SELDESTI09.Label1 <> "OK" Then GoTo 99
   '
   ENVIOMASIVO% = 1
   TODOS% = 0: SELSTRING$ = String$(SELDESTI09.List5.ListCount, "0")
   If SELDESTI09.Check1.Value = 1 Then
       SELSTRING$ = String$(SELDESTI09.List5.ListCount, "1")
     Else
       For KKI% = 1 To SELDESTI09.List5.ListCount
         If SELDESTI09.List5.Selected(KKI% - 1) = True Then Mid$(SELSTRING$, KKI%) = "1"
       Next KKI%
   End If
   '
   Dim CTACSEL As ADODB.Recordset
   Set CTASEL = New ADODB.Recordset
   SQLX$ = "SELECT * FROM BSEGCOM WITH(NOLOCK) ORDER BY CODI_CLI ASC"
   Set CTASEL = FLEXCONN.Execute(FILTSQL$(SQLX$))
   '
   YAENVIOELPRIMER% = 0
   With CTASEL
      Do While Not .EOF
        PERLIST$ = SAFETEXT$(!GRUPOS)
        For KKI% = 1 To Len(PERLIST$)
           PERLI% = Asc(Mid$(PERLIST$, KKI%, 1))
           If PERLI% > 0 Then
             If Mid$(SELSTRING$, PERLI%, 1) = "1" Then
               CODCLISE$ = SAFETEXT$(!CODI_CLI)
               CANMAI% = 0
               Call PREPARAMAIL(CODCLISE$, CANMAI%)
               '
               If CANMAI% > 0 Then
                 If CODCLISE$ <> "" Then
                    TTTT$ = "Envío Mail Masivo (#" & TRIM$(Label13) & "#) " & Format(Now, "dd/mm/yy - hh:mm") + CRLF$ + CRLF$
                    Call FISECOM09.ANOTAHISTORIAL(TTTT$, CODCLISE$)
                 End If
               End If
               '
               CONDI$ = "NUME_MAIL = " & Label13
               If YAENVIOELPRIMER% = 0 Then
                    YAENVIOELPRIMER% = 1
                    Call ACTUREGISTRO("BANDENVI", "Codi_Cli", CONDI$, "VARIOS", RAFE&, "NOMESS")
                    Call ACTUREGISTRO("BANDENVI", "PARA_MAIL", CONDI$, "VARIOS", RAFE&, "NOMESS")
                    Call ACTUREGISTRO("BANDENVI", "ASUN_MAIL", CONDI$, Left$(Text20, 256), RAFE&, "NOMESS")
                    Call ACTUREGISTRO("BANDENVI", "DREL_MAIL", CONDI$, "", RAFE&, "NOMESS")
                    Call ACTUREGISTRO("BANDENVI", "CCOP_MAIL", CONDI$, "", RAFE&, "NOMESS")
                    Call ACTUREGISTRO("BANDENVI", "TEXT_MAIL", CONDI$, RESTODELMAIL$, RAFE&, "NOMESS")
                       SOLICO% = Check1.Value
                    Call ACTUREGISTRO("BANDENVI", "SOLI_CONF", CONDI$, SOLICO%, RAFE&, "NOMESS")
                    Call ACTUREGISTRO("BANDENVI", "CONF_RECEP", CONDI$, FETEXCOR1$(FECHANUM("01/01/80")), RAFE&, "NOMESS")
                       SOLENVI = FETEXCOR1$(HOY(HOS$)) + " " + Time$
                    Call ACTUREGISTRO("BANDENVI", "FHOR_SOLENVI", CONDI$, SOLENVI, RAFE&, "NOMESS")
                    Call ACTUREGISTRO("BANDENVI", "FREC_REPI", CONDI$, 0, RAFE&, "NOMESS")
                    Call ACTUREGISTRO("BANDENVI", "FEFIN_REPI", CONDI$, FETEXCOR1$(FECHANUM(Text7)), RAFE&, "NOMESS")
                    Call ACTUREGISTRO("BANDENVI", "FHOR_ULTENVI", CONDI$, FETEXCOR1$(FECHANUM("01/01/80")), RAFE&, "NOMESS")
               End If
               CAENVI& = 1 + XVALO(VALOBADA("BANDENVI", "CANT_ENVI", CONDI$, "NOMESS"))
               Call ACTUREGISTRO("BANDENVI", "CANT_ENVI", CONDI$, CAENVI&, RAFE&, "NOMESS")

             End If
           End If
        Next KKI%
        '
      .MoveNext
      Loop
   End With
   '
   If ENVIOMASIVO% > 0 Then
      Unload Me
   End If
   '
99 Command2.Enabled = True
   ENVIOMASIVO% = 0
   '
End Sub

Sub PREPARAMAIL(CODCLISE$, CANMAI%)
    '
        'ASIGNO EL CODIGO DEL CLIENTE AL FORMULARIO
        FISECOM09.Label26 = TRIM$(CODCLISE$)
        '
        'RECORRO LOS CONTACTOS
        CANMAI% = 0
        For i% = 0 To FISECOM09.Text6.Count - 1
           With FISECOM09
              PERSO$ = TRIM$(.Text6(i%))
              MAILX$ = TRIM$(.Text8(i%))
              
              If PERSO$ <> "" And MAILX$ <> "" Then
                 'PRESENTO LA INFORMACION EN PANTALLA DEL MAIL Y NOMBRE
                 SHOWMAIL09.Label26 = CODCLISE$
                 SHOWMAIL09.Label14 = PERSO$
                 SHOWMAIL09.Label16 = MAILX$
                 Call ARMADESTI(CODCLISE$, PERSO$)
                 Call Command1_Click ' envia mail
                 CANMAI% = CANMAI% + 1
              End If
           End With
        Next i%
        '
        If CANMAI% < 1 Then
            PERSO$ = FISECOM09.Text4
            MAILX$ = FISECOM09.Text5
            If MAILX$ <> "" Then
               'PRESENTO LA INFORMACION EN PANTALLA DEL MAIL Y NOMBRE
               SHOWMAIL09.Label26 = CODCLISE$
               SHOWMAIL09.Label14 = PERSO$
               SHOWMAIL09.Label16 = MAILX$
               Call ARMADESTI(CODCLISE$, PERSO$)
               Call Command1_Click ' envia mail
               CANMAI% = CANMAI% + 1
            End If
        End If
        
End Sub
'
Private Sub Command18_Click()
     '
     Command18.Enabled = False
     If DERACCE%("ANUMAIPE", "Anulacion de Mail Pendiente de Envio") < 2 Then GoTo 99
     '
     FREPI% = XVALO(Label12)
     If XVALO(FREPI%) < 1 Then
        OPX1% = COMALTER%("¿Confirma la Anulacion de este Mail?\Esta Transaccion no Podra Deshacerse\\No, Conservar\Si, Anular Mail")
        If OPX1% = 2 Then
            NUMAI& = XVALO(Label13)
            CONDI$ = "NUME_MAIL = " & NUMAI&
            Call BORRAREGISTROS("BANDENVI", CONDI$, REG&)
            Call BORRAREGISTROS("ADJUNMAI", CONDI$, REG&)
            Call COMUNI("Anulación Completa")
            Call CIERRARCH("*.*")
            Hide
        End If
     Else
        OPX2% = COMALTER%("Este Mail ya fue Enviado " + Str$(FREPI%) + " Veces\¿Desea Cancelar la Repeticion?\\No, Mantener\Si, Cancelar")
        If OPX2% = 2 Then
           NUMAI& = XVALO(Label13)
           CONDI$ = "NUME_MAIL = " & NUMAI&
           Call ACTUREGISTRO("BANDENVI", "Frec_Repi", CONDI$, 0, REG&, "NOMESS")
           Call COMUNI("Cancelacion Completa")
           Call CIERRARCH("*.*")
           Hide
        End If
     End If
99   Command18.Enabled = True
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   PREPAMAI09.Show 1
   If PREPAMAI09.Label1 = "OK" Then
      PERSODES$ = PREPAMAI09.Combo1.TEXT
      FISECOM09.ORIMAI = "SHOWMAIL09"
      For KKI% = 1 To FISECOM09.Text6.Count
         If TRIM$(FISECOM09.Text6(KKI% - 1)) <> "" Then
            If FISECOM09.Text6(KKI% - 1) = PREPAMAI09.Combo1.TEXT Then
               Call FISECOM09.Text6_DblClick(KKI% - 1)
               GoTo 80
            End If
         End If
      Next KKI%
      Call FISECOM09.Text5_DblClick
      '
80    PREPAMAI09.Label1 = ""
   End If
   Command3.Enabled = True
End Sub

Private Sub Command7_Click()
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Form_Activate()
   If ATENTI% > 0 Then
     Left = 0
     Top = 0
     Call CENTRAFORM(Me)
     ATENTI% = 0
     On Error Resume Next
     Text2.SetFocus
     On Error GoTo 0
   End If
End Sub

Private Sub Form_Load()
   '
   Call CENTRAFORM(Me)
   FEX = DIPOST(DIPOST(HOY(HOS$)))
   Text12 = FECHATEX$(FEX)
   '
End Sub
'
Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Cancel = 0
End Sub

Private Sub Label13_Change()
    NUMAI& = XVALO(Label13)
    CONDI$ = "NUME_MAIL = " & NUMAI&
    If CantRegistros("BANDENVI", CONDI$, "NOMESS") > 0 Then
          Caption = "Muestra Mensaje de Correo Electronico #" & TRIM$(Str$(NUMAI&))
          Label26 = TRIM$(VALOBADA("BANDENVI", "Codi_Cli", CONDI$, "NOMESS"))
          Label14 = TRIM$(VALOBADA("BANDENVI", "PARA_MAIL", CONDI$, "NOMESS"))
          Label15 = TRIM$(VALOBADA("BANDENVI", "ASUN_MAIL", CONDI$, "NOMESS"))
          Label16 = TRIM$(VALOBADA("BANDENVI", "DREL_MAIL", CONDI$, "NOMESS"))
          Label17 = TRIM$(VALOBADA("BANDENVI", "CCOP_MAIL", CONDI$, "NOMESS"))
          Text2 = VALOBADA("BANDENVI", "TEXT_MAIL", CONDI$, "NOMESS")
          Check1.Value = XVALO(VALOBADA("BANDENVI", "SOLI_CONF", CONDI$, "NOMESS"))
          Label18 = TRIM$(Format(VALOBADA("BANDENVI", "CONF_RECEP", CONDI$, "NOMESS"), "dd/mm/yy - hh:mm"))
             If Left$(Label18, 8) = "01/01/80" Then Label18 = ""
          '   SOLENVI = FETEXCOR1$(HOY(HOS$)) + " " + Time$
          SOLENVI$ = TRIM$(Format(VALOBADA("BANDENVI", "FHOR_SOLENVI", CONDI$, "NOMESS"), "dd/mm/yy - hh:mm"))
          Label4 = TRIM$(Left$(SOLENVI$, 8)): Label19 = TRIM$(Mid$(SOLENVI$, 12, 5))
          '
          Text6 = TRIM$(VALOBADA("BANDENVI", "FREC_REPI", CONDI$, "NOMESS"))
          Check2.Value = 0: If XVALO(Text6) > 0 Then Check2.Value = 1
          Text7 = FECHATEX$(CVINT(VALOBADA("BANDENVI", "FEFIN_REPI", CONDI$, "NOMESS")))
          Label2 = TRIM$(Format(VALOBADA("BANDENVI", "FHOR_ULTENVI", CONDI$, "NOMESS"), "dd/mm/yy - hh:mm"))
             If Left$(Label2, 8) = "01/01/80" Then Label2 = ""
          Label12 = XVALO(VALOBADA("BANDENVI", "CANT_ENVI", CONDI$, "NOMESS"))
          Call CARGALISTA(List3, "ADJUNMAI", "NORD_MAIL/F3; FILE_NAME/A128", CONDI$, "NOMESS")
          '
          USUPRE% = XVALO(VALOBADA("BANDENVI", "USUA_PREPA", CONDI$, "NOMESS"))
          Label20 = USERNAM$(USUPRE% Mod 100)
          FHORPRE = VALOBADA("BANDENVI", "FHOR_PREPA", CONDI$, "NOMESS")
          Label10 = FORMATNUM$(FHORPRE, "D20")
       Else
          Call MENSERR(24, "Imposible Abrir Mensaje Enviado Numero '" & TRIM$(Str$(NUMAI&)) & "'")
   End If
End Sub

Private Sub Label4_DblClick()
   Label4 = ""
   Text1 = ""
End Sub

Private Sub List3_DblClick()
   '
   On Error Resume Next
   MkDir LUCOM$ + "ARADJUN"
   On Error GoTo 0
   RUTAR$ = LUCOM$ + "ARADJUN\"
   '
   NUMAI& = XVALO(Label13)
   NORMAI& = XVALO(Left$(List3.TEXT, 3))
   CONDI$ = "NUME_MAIL = " & NUMAI& & " AND NORD_MAIL = " & NORMAI&
   FINAM$ = VALOBADA("ADJUNMAI", "FILE_NAME", CONDI$, "NOMESS")
   XXXX$ = VALOBADA("ADJUNMAI", "FILE_DATA", CONDI$, "NOMESS")
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
   If ASUNTX$ = "" Then ASUNTX$ = Control$("SEGUICOM", "ASUNTO")
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
   If DESTI$ = EMPRE$ Then
         TTYY$ = TTYY$ + "Sres." + Chr$(13) + Chr$(10)
      Else
         TTYY$ = TTYY$ + "Sr./Sra." + Chr$(13) + Chr$(10)
         TTYY$ = TTYY$ + DESTI$ + Chr$(13) + Chr$(10)
   End If
   '
   TTYY$ = TTYY$ + TTXX$ + CRLF + CRLF + RESTODELMAIL$
   Text2 = TTYY$
   '
End Sub


