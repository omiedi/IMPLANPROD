VERSION 5.00
Begin VB.Form PREIMPRE 
   Caption         =   "Opciones de Impresión"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   7230
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
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
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   2040
         TabIndex        =   10
         Text            =   "Vertical"
         Top             =   3120
         Width           =   2055
      End
      Begin VB.TextBox txtEncabezado1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   525
         Left            =   2040
         TabIndex        =   5
         Text            =   "Reporte 1"
         Top             =   240
         Width           =   4935
      End
      Begin VB.TextBox txtEncabezado2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   645
         Left            =   2040
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   960
         Width           =   4935
      End
      Begin VB.TextBox txtPiePagina 
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
         Left            =   2040
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1920
         Width           =   4935
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2040
         TabIndex        =   2
         Text            =   "Vertical"
         Top             =   2640
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Imprimir"
         Height          =   855
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   2640
         Width           =   2295
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label6"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   1320
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Seleccionar Impresora"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   840
         TabIndex        =   11
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Encabezado"
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
         Left            =   960
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Subtítulo"
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
         Left            =   1200
         TabIndex        =   8
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Pie de Página"
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
         Left            =   720
         TabIndex        =   7
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Orientación del papel:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   6
         Top             =   2520
         Width           =   1695
      End
   End
End
Attribute VB_Name = "PREIMPRE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objPrinterFlex As PrinterFlex
Private Sub Command1_Click()
Command1.Enabled = False
    'LLAMO AL EVENTO CLICK PARA QUE ACTUALICE la orientación de la impresión
     Call Combo1_Click ' COMBO DE ORIENTACION DE IMPRESION
 
    With objPrinterFlex
      'Asignamos los valores de los encabezados, el pie de página, el color_
      'del texto y el tamaño de la fuente
        
        'texto de los encabezdos y el pir de pagina
        .TextEncabezado1 = txtEncabezado1.TEXT
        .TextEncabezado2 = txtEncabezado2.TEXT
        .TextPiePagina = txtPiePagina
        
        'Colores de la fuentes
        .ColorPiePagina = txtPiePagina.ForeColor
        .ColorEncabezado1 = txtEncabezado1.ForeColor
        .ColorEncabezado2 = txtEncabezado2.ForeColor
        
        'Tamaños de las fuentes
        .SizeEncabezado1 = txtEncabezado1.FontSize
        .SizeEncabezado2 = txtEncabezado2.FontSize
        .SizePiePagina = txtPiePagina.FontSize
        .AjustarColumnas = True
      
        'Imprimimos pasando el nombre del FlexGrid a imprimir
        .ImprimirFlexGrid GRILLA_IMPRESION
    End With

Command1.Enabled = True
End Sub

Private Sub Form_Load()

    Combo2.Clear
    Dim PRX As Printer
    On Error Resume Next
    Combo2.TEXT = Printer.DeviceName
    On Error GoTo 0
    For Each PRX In Printers
      Combo2.AddItem PRX.DeviceName
      If PRX.DeviceName = IMPRE$ Then
        Combo2.TEXT = IMPRE$
      End If
    Next
    '
   'Inicializamos la variable objPrinterFlex
    Set objPrinterFlex = New PrinterFlex

End Sub
'Indica la orientación de la impresión
Private Sub Combo1_Click()

    If Combo1.ListIndex = 0 Then
       objPrinterFlex.Orientacion = Vertical
    Else
       objPrinterFlex.Orientacion = Horizontal
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Removemos
    Set objPrinterFlex = Nothing
    End Sub

