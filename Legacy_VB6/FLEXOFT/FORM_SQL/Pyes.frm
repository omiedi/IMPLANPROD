VERSION 5.00
Begin VB.Form Pyes 
   Caption         =   "Form1"
   ClientHeight    =   7140
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15420
   LinkTopic       =   "Form1"
   ScaleHeight     =   7140
   ScaleWidth      =   15420
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   13440
      TabIndex        =   11
      Top             =   0
      Width           =   1215
   End
   Begin VB.TextBox TXT1 
      Height          =   495
      Index           =   4
      Left            =   960
      TabIndex        =   10
      Text            =   "Text1"
      ToolTipText     =   "stringEntidadExcepcion"
      Top             =   1920
      Width           =   14000
   End
   Begin VB.TextBox TXT1 
      Height          =   495
      Index           =   3
      Left            =   960
      TabIndex        =   9
      Text            =   "Text1"
      ToolTipText     =   "apellidoVendedor"
      Top             =   1440
      Width           =   14000
   End
   Begin VB.TextBox TXT1 
      Height          =   495
      Index           =   2
      Left            =   960
      TabIndex        =   8
      Text            =   "Text1"
      ToolTipText     =   "nombreVendedor"
      Top             =   960
      Width           =   14000
   End
   Begin VB.TextBox TXT1 
      Height          =   495
      Index           =   1
      Left            =   960
      TabIndex        =   3
      Text            =   "Text1"
      ToolTipText     =   "entidad1"
      Top             =   480
      Width           =   14000
   End
   Begin VB.TextBox TXT1 
      Height          =   495
      Index           =   0
      Left            =   960
      TabIndex        =   2
      Text            =   "Text1"
      ToolTipText     =   "codigo_interno_VEND_ELEG"
      Top             =   0
      Width           =   11355
   End
   Begin VB.TextBox Text19 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4305
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   2520
      Width           =   15315
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Pyes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_Change()

End Sub

Private Sub Command1_Click()

'1.-SE DEBERA PONER EN EL TOOLTIPTEXT EL TEXTO LITERAL QUE FIGURA EN EL ARCHIVO AL CUAL SE REEMPLAZARA
'2.-EN LAS CAJAS DE TEXTOS ESCRIBIR EL TEXTO REEMPLAZANTE PARA CADA TEXTO DEL TOOLTIPTEXT INGRESADO EN 1

TEXTOENPANTALLA$ = ""
Text19 = ""
J& = 0
    NARCHI1% = FILEOPEN%("C:\Users\Omar\Desktop\UNIFICACION DE VENDEDORES\INSTRUCCIONES DE UNIFICACION DE VENDEDORES_ORIGINAL_3.txt", 1, 1024) 'ABRE EL ARCHIVO DE LECTURA
'    NARCHI2% = FILEOPEN%(CARPETOK$ + "consulta_cae_WSBFE.txt", 2, 128) ' ABRE EL ARCHIVO DE ESCRITURA
    Dim CONTROLX As CONTROL
    Do While Not EOF(NARCHI1%) ' RECORRE
       LINEA = LINEA + 1
       J& = 0
       Line Input #NARCHI1%, XX$ ' LEE CADA REGISTRO
          For Each CONTROLX In Me.Controls ' RECORRO LOS CONTROLES DEL FORM ACTIVO
            If TypeOf CONTROLX Is TextBox Then
              If J& <= TXT1.UBound Then
                If InStr(XX$, TXT1(J&).ToolTipText) > 0 Then
                  XX$ = REPLACAR(XX$, TXT1(J&).ToolTipText, TXT1(J&).TEXT)
                End If
              Else
                Exit For
              End If
              J& = J& + 1
            End If
            TEXTOENPANTALLA$ = TEXTOENPANTALLA$ & XX$
         Next
         TEXTOENPANTALLA$ = TEXTOENPANTALLA$ & XX$ & vbCrLf
         Debug.Print XX$
         If LINEA > 100 Then
         A = 1
         End If
     Loop
     
     
     Close #NARCHI1%
     Text19 = TEXTOENPANTALLA$
     

End Sub

