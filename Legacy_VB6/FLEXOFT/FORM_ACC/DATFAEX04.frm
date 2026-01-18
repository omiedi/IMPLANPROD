VERSION 5.00
Begin VB.Form DATFAEX04 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Datos Complementarios de Factura de Exportacion"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo6 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   34
      ToolTipText     =   "Click Para Desplegar"
      Top             =   3240
      Width           =   4905
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
      Height          =   615
      Left            =   6600
      Picture         =   "DATFAEX04.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Aprueba la Información"
      Top             =   7000
      Width           =   795
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00404080&
      Height          =   6945
      Left            =   6480
      ScaleHeight     =   6885
      ScaleWidth      =   1635
      TabIndex        =   31
      Top             =   0
      Width           =   1695
      Begin VB.PictureBox Picture16 
         Height          =   7000
         Left            =   0
         ScaleHeight     =   6945
         ScaleWidth      =   1020
         TabIndex        =   36
         Top             =   0
         Visible         =   0   'False
         Width           =   1080
         Begin VB.Image Image3 
            Height          =   7005
            Left            =   0
            Picture         =   "DATFAEX04.frx":030A
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.CommandButton Command13 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   150
         Picture         =   "DATFAEX04.frx":26A2
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   1080
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   32
         Top             =   6120
         Width           =   645
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   33
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command5 
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
         Height          =   690
         Left            =   150
         Picture         =   "DATFAEX04.frx":29AC
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Salir - Cancelar"
         Top             =   240
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -255
         Picture         =   "DATFAEX04.frx":2AF6
         Top             =   6480
         Width           =   1515
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Permisos de Embarque"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2235
      Left            =   240
      TabIndex        =   12
      Top             =   5400
      Width           =   6120
      Begin VB.CommandButton Command3 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5520
         TabIndex        =   30
         ToolTipText     =   "Elimina el Permiso Seleccionado"
         Top             =   1680
         Width           =   375
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
         Height          =   375
         Left            =   5040
         Picture         =   "DATFAEX04.frx":4A18
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Ingresa el Permiso en la Ventana Inferior"
         Top             =   1680
         Width           =   375
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
         Height          =   360
         Left            =   240
         MaxLength       =   16
         TabIndex        =   4
         ToolTipText     =   "Ingreso el Permiso"
         Top             =   1680
         Width           =   4665
      End
      Begin VB.ListBox List2 
         BackColor       =   &H00FFFFFF&
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
         ItemData        =   "DATFAEX04.frx":4D22
         Left            =   240
         List            =   "DATFAEX04.frx":4D24
         MousePointer    =   1  'Arrow
         Sorted          =   -1  'True
         TabIndex        =   13
         Top             =   600
         Width           =   5625
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Lista de Permisos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   5625
      End
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   3
      ToolTipText     =   "Click Para Desplegar"
      Top             =   2850
      Width           =   4905
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   2
      ToolTipText     =   "Click Para Desplegar"
      Top             =   2430
      Width           =   4905
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Comprobantes Asociados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1755
      Left            =   240
      TabIndex        =   9
      Top             =   3600
      Width           =   6120
      Begin VB.Frame Frame4 
         Height          =   1440
         Left            =   5520
         TabIndex        =   21
         Top             =   240
         Width           =   15
      End
      Begin VB.Frame Frame14 
         Height          =   1425
         Left            =   3960
         TabIndex        =   20
         Top             =   240
         Width           =   15
      End
      Begin VB.Frame Frame3 
         Caption         =   "                "
         Height          =   1440
         Left            =   2520
         TabIndex        =   19
         Top             =   240
         Width           =   15
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1140
         ItemData        =   "DATFAEX04.frx":4D26
         Left            =   240
         List            =   "DATFAEX04.frx":4D28
         MousePointer    =   1  'Arrow
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   480
         Width           =   5625
      End
      Begin VB.Label Label25 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo Comprobante             Punto Venta      Número"
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
         TabIndex        =   11
         ToolTipText     =   "Doble click Para Editar"
         Top             =   240
         Width           =   5625
      End
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Click Para Desplegar"
      Top             =   2000
      Width           =   4905
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Click Para Desplegar"
      Top             =   1600
      Width           =   4905
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
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
      Height          =   1335
      Left            =   240
      TabIndex        =   23
      Top             =   120
      Width           =   6135
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
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
         Left            =   0
         TabIndex        =   29
         Top             =   270
         Width           =   750
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000016&
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
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   480
         TabIndex        =   28
         Top             =   525
         Width           =   4140
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
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
         Left            =   4440
         TabIndex        =   27
         Top             =   270
         Width           =   750
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
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
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   4920
         TabIndex        =   26
         Top             =   525
         Width           =   1020
      End
      Begin VB.Label Label14 
         BackColor       =   &H80000016&
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
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   3840
         TabIndex        =   25
         Top             =   930
         Width           =   2100
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Comprobante:"
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
         Left            =   2520
         TabIndex        =   24
         Top             =   1000
         Width           =   1230
      End
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Idioma:"
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
      Left            =   670
      TabIndex        =   35
      Top             =   3300
      Width           =   885
   End
   Begin VB.Label Label17 
      Caption         =   "Label17"
      Height          =   255
      Left            =   0
      TabIndex        =   22
      Top             =   7320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Incoterms:"
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
      Left            =   400
      TabIndex        =   17
      Top             =   2955
      Width           =   1485
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Cuit Pais:"
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
      Left            =   490
      TabIndex        =   16
      Top             =   2520
      Width           =   1005
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Destino:"
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
      Left            =   600
      TabIndex        =   15
      Top             =   2085
      Width           =   885
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Exp:"
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
      Left            =   480
      TabIndex        =   8
      Top             =   1680
      Width           =   1485
   End
   Begin VB.Menu archivo 
      Caption         =   "Archivo"
      Begin VB.Menu grabar 
         Caption         =   "Aprobar"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "DATFAEX04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub BLANFORMU()
  
   Label9 = ""
   Label10 = ""
   Label14 = ""
   List1.Clear
   List2.Clear
   Text1 = ""
End Sub

Private Sub Command1_Click()
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   '
   'VALIDACIONES
   If TRIM$(Combo1.TEXT) = "" Then
      Call COMUNI("Falta Ingresar Tipo de Exportación")
      GoTo 99
   End If
   '
   If Left$((Combo1.TEXT), 1) <> "1" Then
      If List2.ListCount > 0 And Left$(Label14, 2) = "FC" Then
        Call COMUNI("Tipo de Exportación (" & TIPO_EXPO$ & ") No Válido Con Ingreso de Exportación")
        GoTo 99
      End If
   End If
   '
   If TRIM$(Combo2.TEXT) = "" Then
      Call COMUNI("Falta Ingresar Destino")
      GoTo 99
   End If
   '
   If TRIM$(Combo5.TEXT) = "" Then
      Call COMUNI("Falta Ingresar Incoterms")
      GoTo 99
   End If
   '
   If TRIM$(Combo6.TEXT) = "" Then
      Call COMUNI("Falta Ingresar Idioma")
      GoTo 99
   End If
   'FIN VALIDACIONES
   '
   If Left$(Label14, 2) = "FC" And List1.ListCount > 0 Then
      Call COMUNI("Comprobantes Asociados Solo para Nota de Crédito y Nota de Débito")
      GoTo 99
   End If
   '
   If TRIM$(Combo1.TEXT) = "" Then
      Call COMUNI("Falta Ingresar Idioma")
      GoTo 99
   End If
   '
   Label17 = "OK"
   Hide
   
99 Command2.Enabled = True

End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    'borra el permiso seleccionado
    If List2.ListIndex >= 0 Then
      List2.RemoveItem (List2.ListIndex)
      Text1.SetFocus
    End If
    '
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   If TRIM$(Text1) <> "" Then
      List2.AddItem Text1
      Text1 = ""
   End If
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
   Label17 = ""
   Screen.MousePointer = 1
   Command5.Enabled = True
   Hide
End Sub

Private Sub Form_Load()
  If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      On Error Resume Next
      Picture16.Visible = True
      Picture16.ZOrder 0
      Me.BackColor = &HFCD7B6
      Me.Frame14.BackColor = &HFCD7B6
      Me.Frame1.BackColor = &HFCD7B6
      Me.Frame2.BackColor = &HFCD7B6
      Me.Frame3.BackColor = &HFCD7B6
      Me.Frame4.BackColor = &HFCD7B6
      Me.Frame5.BackColor = &HFCD7B6
            
      MARCOBARRA.Top = 6600
      
      Set MARCOBARRA.Container = Me.Picture16
      MARCOBARRA.ZOrder 0
      On Error GoTo 0
  End If

  Label17 = ""
  'Call BLANFORMU
  Screen.MousePointer = 11
  '
  TIFACT$ = "FC-"
  If FORFACCS.Label7.Caption = "N.Débito:" Then TIFACT$ = "ND-"
  If FORFACCS.Label7.Caption = "N.Crédito:" Then TIFACT$ = "NC-"
  Label9 = RASOCLI$(XVALO(FORFACCS.Text1))
  Label10 = FORFACCS.Text9
  NUM$ = FORFACCS.Text11
  While Len(NUM$) < 8: NUM$ = "0" & NUM$: Wend
  Label14 = TIFACT$ & FORFACCS.Label12 & "-" & NUM$
  '
  Call CARGACOMBO1(Combo1)
  Call CARGACOMBO2(Combo2)
  Call CARGACOMBO4(Combo4)
  Call CARGACOMBO5(Combo5)
  Call CARGACOMBO6(Combo6)
  Label12 = CONTROL("FACTURA", "PUVEFAEX")
  Combo1.TEXT = Combo1.List(1)
  Combo5.TEXT = Combo5.List(4)
  '
  'CARGO TABLA DE MOTIVOS SI NO ESTABA CARGADA PREVIAMENTE
  If EXISTE(LUDAT$ + "\TATICOMP.RFS") > 0 Then
     If ULTREG&("TATICOMP") < 3 Then
            Call GRAREG("TATICOMP", Chr$(19) + "Factura", 1)
            Call GRAREG("TATICOMP", Chr$(20) + "Nota de Credito", 2)
            Call GRAREG("TATICOMP", Chr$(21) + "Nota de Debito", 3)
            Call CIERRARCH("TATICOMP")
     End If
  End If
  Screen.MousePointer = 1
End Sub
Sub CARGACOMBO2(COMBOX As ComboBox)
    '
    COMBOX.AddItem ""
    COMBOX.AddItem "101    BURKINA FASO"
    COMBOX.AddItem "102    ARGELIA"
    COMBOX.AddItem "103    BOTSWANA"
    COMBOX.AddItem "104    BURUNDI"
    COMBOX.AddItem "105    CAMERUN"
    COMBOX.AddItem "107    REP. CENTROAFRICANA."
    COMBOX.AddItem "108    CONGO"
    COMBOX.AddItem "109    REP.DEMOCRAT.DEL CONGO EX ZAIRE"
    COMBOX.AddItem "110    COSTA DE MARFIL"
    COMBOX.AddItem "111    CHAD"
    COMBOX.AddItem "112    BENIN"
    COMBOX.AddItem "113    EGIPTO"
    COMBOX.AddItem "115    GABON"
    COMBOX.AddItem "116    GAMBIA"
    COMBOX.AddItem "117    GHANA"
    COMBOX.AddItem "118    GUINEA"
    COMBOX.AddItem "119    GUINEA ECUATORIAL"
    COMBOX.AddItem "120    KENYA"
    COMBOX.AddItem "121    LESOTHO"
    COMBOX.AddItem "122    LIBERIA"
    COMBOX.AddItem "123    LIBIA"
    COMBOX.AddItem "124    MADAGASCAR"
    COMBOX.AddItem "125    MALAWI"
    COMBOX.AddItem "126    MALI"
    COMBOX.AddItem "127    MARRUECOS"
    COMBOX.AddItem "128    MAURICIO , ISLAS"
    COMBOX.AddItem "129    MAURITANIA"
    COMBOX.AddItem "130    NIGER"
    COMBOX.AddItem "131    NIGERIA"
    COMBOX.AddItem "132    ZIMBABWE"
    COMBOX.AddItem "133    RWANDA"
    COMBOX.AddItem "134    SENEGAL"
    COMBOX.AddItem "135    SIERRA LEONA"
    COMBOX.AddItem "136    SOMALIA"
    COMBOX.AddItem "137    SWAZILANDIA"
    COMBOX.AddItem "138    SUDAN"
    COMBOX.AddItem "139    TANZANIA"
    COMBOX.AddItem "140    TOGO"
    COMBOX.AddItem "141    TUNEZ"
    COMBOX.AddItem "142    UGANDA"
    COMBOX.AddItem "144    ZAMBIA"
    COMBOX.AddItem "145    TERRIT.VINCULADOS AL R UNIDO"
    COMBOX.AddItem "146    TERRIT.VINCULADOS A ESPAÑA"
    COMBOX.AddItem "147    TERRIT.VINCULADOS A FRANCIA"
    COMBOX.AddItem "149    ANGOLA"
    COMBOX.AddItem "150    CABO VERDE"
    COMBOX.AddItem "151    MOZAMBIQUE"
    COMBOX.AddItem "152    SEYCHELLES"
    COMBOX.AddItem "153    DJIBOUTI"
    COMBOX.AddItem "155    COMORAS"
    COMBOX.AddItem "156    GUINEA BISSAU"
    COMBOX.AddItem "157    STO.TOME Y PRINCIPE"
    COMBOX.AddItem "158    NAMIBIA"
    COMBOX.AddItem "159    SUDAFRICA"
    COMBOX.AddItem "160    ERITREA"
    COMBOX.AddItem "161    ETIOPIA"
    COMBOX.AddItem "197    RESTO (AFRICA)"
    COMBOX.AddItem "198    INDETERMINADO (AFRICA)"
    COMBOX.AddItem "200    ARGENTINA"
    COMBOX.AddItem "201    BARBADOS"
    COMBOX.AddItem "202    BOLIVIA"
    COMBOX.AddItem "203    BRASIL"
    COMBOX.AddItem "204    CANADA"
    COMBOX.AddItem "205    COLOMBIA"
    COMBOX.AddItem "206    COSTA RICA"
    COMBOX.AddItem "207    CUBA"
    COMBOX.AddItem "208    CHILE"
    COMBOX.AddItem "209    REPÚBLICA DOMINICANA"
    COMBOX.AddItem "210    ECUADOR"
    COMBOX.AddItem "211    EL SALVADOR"
    COMBOX.AddItem "212    ESTADOS UNIDOS"
    COMBOX.AddItem "213    GUATEMALA"
    COMBOX.AddItem "214    GUYANA"
    COMBOX.AddItem "215    HAITI"
    COMBOX.AddItem "216    HONDURAS"
    COMBOX.AddItem "217    JAMAICA"
    COMBOX.AddItem "218    MEXICO"
    COMBOX.AddItem "219    NICARAGUA"
    COMBOX.AddItem "220    PANAMA"
    COMBOX.AddItem "221    PARAGUAY"
    COMBOX.AddItem "222    PERU"
    COMBOX.AddItem "223    Puerto RICO"
    COMBOX.AddItem "224    TRINIDAD Y TOBAGO"
    COMBOX.AddItem "225    URUGUAY"
    COMBOX.AddItem "226    VENEZUELA"
    COMBOX.AddItem "227    TERRIT.VINCULADO AL R.UNIDO"
    COMBOX.AddItem "228    TER.VINCULADOS A DINAMARCA"
    COMBOX.AddItem "229    TERRIT.VINCULADOS A FRANCIA AMERIC."
    COMBOX.AddItem "230    TERRIT.HOLANDESES"
    COMBOX.AddItem "231    TER.VINCULADOS A ESTADOS UNIDOS"
    COMBOX.AddItem "232    SURINAME"
    COMBOX.AddItem "233    DOMINICA"
    COMBOX.AddItem "234    SANTA LUCIA"
    COMBOX.AddItem "235    SAN VICENTE Y LAS GRANADINAS"
    COMBOX.AddItem "236    BELICE"
    COMBOX.AddItem "237    ANTIGUA Y BARBUDA"
    COMBOX.AddItem "238    S.CRISTOBAL Y NEVIS"
    COMBOX.AddItem "239    BAHAMAS"
    COMBOX.AddItem "240    GRENADA"
    COMBOX.AddItem "241    ANTILLAS HOLANDESAS"
    COMBOX.AddItem "250    AAE Tierra del Fuego - ARGENTINA"
    COMBOX.AddItem "251    ZF La Plata - ARGENTINA"
    COMBOX.AddItem "252    ZF Justo Daract - ARGENTINA"
    COMBOX.AddItem "253    ZF Río Gallegos - ARGENTINA"
    COMBOX.AddItem "254    ISLAS Malvinas - ARGENTINA"
    COMBOX.AddItem "255    ZF Tucumán - ARGENTINA"
    COMBOX.AddItem "256    ZF Córdoba - ARGENTINA"
    COMBOX.AddItem "257    ZF Mendoza - ARGENTINA"
    COMBOX.AddItem "258    ZF General Pico - ARGENTINA"
    COMBOX.AddItem " 259    ZF Comodoro Rivadavia - ARGENTINA"
    COMBOX.AddItem "260    ZF Iquique"
    COMBOX.AddItem "261    ZF Punta Arenas"
    COMBOX.AddItem "262    ZF Salta - ARGENTINA"
    COMBOX.AddItem "263    ZF Paso de los Libres - ARGENTINA"
    COMBOX.AddItem "264    ZF Puerto Iguazú - ARGENTINA"
    COMBOX.AddItem "265    SECTOR ANTARTICO ARG."
    COMBOX.AddItem "270    ZF Colón - REPÚBLICA DE PANAMÁ"
    COMBOX.AddItem "271    ZF Winner (Sta. C. de la Sierra) - BOLIVIA"
    COMBOX.AddItem "280    ZF Colonia - URUGUAY"
    COMBOX.AddItem "281    ZF Florida - URUGUAY"
    COMBOX.AddItem "283    ZF Zonamerica - URUGUAY"
    COMBOX.AddItem "284    ZF Nueva Helvecia - URUGUAY"
    COMBOX.AddItem "285    ZF Nueva Palmira - URUGUAY"
    COMBOX.AddItem "286    ZF Río Negro - URUGUAY"
    COMBOX.AddItem "287    ZF Rivera - URUGUAY"
    COMBOX.AddItem "288    ZF San José - URUGUAY"
    COMBOX.AddItem "291    ZF Manaos - BRASIL"
    COMBOX.AddItem "295    MAR ARG ZONA ECO.EX"
    COMBOX.AddItem "296    RIOS ARG NAVEG INTER"
    COMBOX.AddItem "297    RESTO AMERICA"
    COMBOX.AddItem "298    INDETERMINADO (AMERICA)"
    COMBOX.AddItem "301    AFGANISTAN"
    COMBOX.AddItem "302    ARABIA SAUDITA"
    COMBOX.AddItem "303    BAHREIN"
    COMBOX.AddItem "304    MYANMAR (EX - BIRMANIA)"
    COMBOX.AddItem "305    BUTAN"
    COMBOX.AddItem "306    CAMBODYA (EX - KAMPUCHE)"
    COMBOX.AddItem "307    SRI LANKA"
    COMBOX.AddItem "308    COREA DEMOCRATICA"
    COMBOX.AddItem "309    COREA REPUBLICANA"
    COMBOX.AddItem "310    CHINA"
    COMBOX.AddItem "312    FILIPINAS"
    COMBOX.AddItem "313    TAIWAN"
    COMBOX.AddItem "315    INDIA"
    COMBOX.AddItem "316    INDONESIA"
    COMBOX.AddItem "317    IRAK"
    COMBOX.AddItem "318    IRAN"
    COMBOX.AddItem "319    ISRAEL"
    COMBOX.AddItem "320    JAPON"
    COMBOX.AddItem "321    JORDANIA"
    COMBOX.AddItem "322    QATAR"
    COMBOX.AddItem "323    KUWAIT"
    COMBOX.AddItem "324    LAOS"
    COMBOX.AddItem "325    LIBANO"
    COMBOX.AddItem "326    MALASIA"
    COMBOX.AddItem "327    MALDIVAS ISLAS"
    COMBOX.AddItem "328    OMAN"
    COMBOX.AddItem "329    MONGOLIA"
    COMBOX.AddItem "330    NEPAL"
    COMBOX.AddItem "331    EMIRATOS ARABES UNIDOS"
    COMBOX.AddItem "332    PAKISTÁN"
    COMBOX.AddItem "333    SINGAPUR"
    COMBOX.AddItem "334    SIRIA"
    COMBOX.AddItem "335    THAILANDIA"
    COMBOX.AddItem "337    VIETNAM"
    COMBOX.AddItem "341    HONG KONG"
    COMBOX.AddItem "344    MACAO"
    COMBOX.AddItem "345    BANGLADESH"
    COMBOX.AddItem "346    BRUNEI"
    COMBOX.AddItem "348    REPUBLICA DE YEMEN"
    COMBOX.AddItem "349    ARMENIA"
    COMBOX.AddItem "350    AZERBAIJAN"
    COMBOX.AddItem "351    GEORGIA"
    COMBOX.AddItem "352    KAZAJSTAN"
    COMBOX.AddItem "353    KIRGUIZISTAN"
    COMBOX.AddItem "354    TAYIKISTAN"
    COMBOX.AddItem "355    TURKMENISTAN"
    COMBOX.AddItem "356    UZBEKISTAN"
    COMBOX.AddItem "357    TERR.AU.PALESTINOS"
    COMBOX.AddItem "397    RESTO DE ASIA"
    COMBOX.AddItem "398    INDET.(ASIA)"
    COMBOX.AddItem "401    ALBANIA"
    COMBOX.AddItem "404    ANDORRA"
    COMBOX.AddItem "405    AUSTRIA"
    COMBOX.AddItem "406    BELGICA"
    COMBOX.AddItem "407    BULGARIA"
    COMBOX.AddItem "409    DINAMARCA"
    COMBOX.AddItem "410    ESPAÑA"
    COMBOX.AddItem "411    FINLANDIA"
    COMBOX.AddItem "412    FRANCIA"
    COMBOX.AddItem "413    GRECIA"
    COMBOX.AddItem "414    HUNGRIA"
    COMBOX.AddItem "415    IRLANDA"
    COMBOX.AddItem "416    ISLANDIA"
    COMBOX.AddItem "417    ITALIA"
    COMBOX.AddItem "418    LIECHTENSTEIN"
    COMBOX.AddItem "419    LUXEMBURGO"
    COMBOX.AddItem "420    MALTA"
    COMBOX.AddItem "421    MONACO"
    COMBOX.AddItem "422    NORUEGA"
    COMBOX.AddItem "423    PAISES BAJOS"
    COMBOX.AddItem "424    POLONIA"
    COMBOX.AddItem "425    PORTUGAL"
    COMBOX.AddItem "426    REINO UNIDO"
    COMBOX.AddItem "427    RUMANIA"
    COMBOX.AddItem "428    SAN MARINO"
    COMBOX.AddItem "429    SUECIA"
    COMBOX.AddItem "430    SUIZA"
    COMBOX.AddItem "431    VATICANO(SANTA SEDE)"
    COMBOX.AddItem "433    POS.BRIT.(EUROPA)"
    COMBOX.AddItem "435    CHIPRE"
    COMBOX.AddItem "436    TURQUIA"
    COMBOX.AddItem "438    ALEMANIA,REP.FED."
    COMBOX.AddItem "439    BIELORRUSIA"
    COMBOX.AddItem "440    ESTONIA"
    COMBOX.AddItem "441    LETONIA"
    COMBOX.AddItem "442    LITUANIA"
    COMBOX.AddItem "443    MOLDAVIA"
    COMBOX.AddItem "444    RUSIA"
    COMBOX.AddItem "445    UCRANIA"
    COMBOX.AddItem "446    BOSNIA HERZEGOVINA"
    COMBOX.AddItem "447    CROACIA"
    COMBOX.AddItem "448    ESLOVAQUIA"
    COMBOX.AddItem "449    ESLOVENIA"
    COMBOX.AddItem "450    MACEDONIA"
    COMBOX.AddItem "451    REP.CHECA"
    COMBOX.AddItem "453    MONTENEGRO"
    COMBOX.AddItem "454    SERBIA"
    COMBOX.AddItem "497    RESTO EUROPA"
    COMBOX.AddItem "498    INDET.(EUROPA)"
    COMBOX.AddItem "501    AUSTRALIA"
    COMBOX.AddItem "503    NAURU"
    COMBOX.AddItem "504    NUEVA ZELANDIA"
    COMBOX.AddItem "505    VANATU"
    COMBOX.AddItem "506    SAMOA OCCIDENTAL"
    COMBOX.AddItem "507    TERRITORIO VINCULADOS A AUSTRALIA"
    COMBOX.AddItem "508    TERRITORIOS VINCULADOS AL R. UNIDO"
    COMBOX.AddItem "509    TERRITORIOS VINCULADOS A FRANCIA"
    COMBOX.AddItem "510    TER VINCULADOS A NUEVA. ZELANDA"
    COMBOX.AddItem "511    TER. VINCULADOS A ESTADOS UNIDOS"
    COMBOX.AddItem "512    FIJI , ISLAS"
    COMBOX.AddItem "513    PAPUA NUEVA GUINEA"
    COMBOX.AddItem "514    KIRIBATI , ISLAS"
    COMBOX.AddItem "515    MICRONESIA , EST.FEDER"
    COMBOX.AddItem "516    PALAU"
    COMBOX.AddItem "517    TUVALU"
    COMBOX.AddItem "518    SALOMON , ISLAS"
    COMBOX.AddItem "519    TONGA"
    COMBOX.AddItem "520    MARSHALL , ISLAS"
    COMBOX.AddItem "521    MARIANAS , ISLAS"
    COMBOX.AddItem "597    RESTO OCEANIA"
    COMBOX.AddItem "598    INDET.(OCEANIA)"
    COMBOX.AddItem "997    RESTO CONTINENTE"
    COMBOX.AddItem "998    INDET.(CONTINENTE)"

End Sub
Sub CARGACOMBO1(COMBOX As ComboBox)
   
    COMBOX.AddItem ""
    COMBOX.AddItem "1  Exportacion definitiva de bienes"
    COMBOX.AddItem "2  Servicios"
    COMBOX.AddItem "4  Otros"
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Label17 = ""
   Screen.MousePointer = 1

End Sub

Private Sub Label11_Click()

End Sub

Private Sub Label5_Click()

End Sub

Private Sub Text10_LostFocus()
   While Len(Text2) < 8: Text2 = "0" & Text2: Wend
End Sub

Sub CARGACOMBO4(COMBOX As ComboBox)
    COMBOX.AddItem ""
    COMBOX.AddItem "50000000016         URUGUAY - Persona Física"
    COMBOX.AddItem "50000000024         PARAGUAY - Persona Física"
    COMBOX.AddItem "50000000032         CHILE - Persona Física"
    COMBOX.AddItem "50000000040         BOLIVIA - Persona Física"
    COMBOX.AddItem "50000000059         BRASIL - Persona Física"
    COMBOX.AddItem "50000001012         BURKINA FASO - Persona Física"
    COMBOX.AddItem "50000001020         ARGELIA - Persona Física"
    COMBOX.AddItem "50000001039         BOTSWANA - Persona Física"
    COMBOX.AddItem "50000001047         BURUNDI - Persona Física"
    COMBOX.AddItem "50000001055         CAMERUN - Persona Física"
    COMBOX.AddItem "50000001071         CENTRO AFRICANO, REP. - Persona Física"
    COMBOX.AddItem "50000001101         COSTA DE MARFIL - Persona Física"
    COMBOX.AddItem "50000001136         EGIPTO - Persona Física"
    COMBOX.AddItem "50000001152         GABON - Persona Física"
    COMBOX.AddItem "50000001160         GAMBIA - Persona Física"
    COMBOX.AddItem "50000001179         GHANA - Persona Física"
    COMBOX.AddItem "50000001187         GUINEA - Persona Física"
    COMBOX.AddItem "50000001195         GUINEA ECUATORIAL - Persona Física"
    COMBOX.AddItem "50000001209         KENIA - Persona Física"
    COMBOX.AddItem "50000001217         LESOTHO - Persona Física"
    COMBOX.AddItem "50000001225         REPUBLICA DE LIBERIA (Estado independiente) - Persona Física"
    COMBOX.AddItem "50000001233         LIBIA - Persona Física"
    COMBOX.AddItem "50000001241         MADAGASCAR - Persona Física"
    COMBOX.AddItem "50000001276         MARRUECOS - Persona Física"
    COMBOX.AddItem "50000001284         REPUBLICA DE MAURICIO - Persona Física"
    COMBOX.AddItem "50000001292         MAURITANIA - Persona Física"
    COMBOX.AddItem "50000001306         NIGER - Persona Física"
    COMBOX.AddItem "50000001314         NIGERIA - Persona Física"
    COMBOX.AddItem "50000001322         ZIMBABWE - Persona Física"
    COMBOX.AddItem "50000001330         RUANDA - Persona Física"
    COMBOX.AddItem "50000001349         SENEGAL - Persona Física"
    COMBOX.AddItem "50000001357         SIERRA LEONA - Persona Física"
    COMBOX.AddItem "50000001365         SOMALIA - Persona Física"
    COMBOX.AddItem "50000001373         REINO DE SWAZILANDIA (Estado independiente) - Persona Física"
    COMBOX.AddItem "50000001381         SUDAN - Persona Física"
    COMBOX.AddItem "50000001403         TOGO - Persona Física"
    COMBOX.AddItem "50000001411         REPUBLICA TUNECINA - Persona Física"
    COMBOX.AddItem "50000001446         ZAMBIA - Persona Física"
    COMBOX.AddItem "50000001454         POS.BRITANICA (AFRICA) - Persona Física"
    COMBOX.AddItem "50000001462         POS.ESPAÑOLA (AFRICA) - Persona Física"
    COMBOX.AddItem "50000001470         POS.FRANCESA (AFRICA) - Persona Física"
    COMBOX.AddItem "50000001489         POS.PORTUGUESA (AFRICA) - Persona Física"
    COMBOX.AddItem "50000001497         REPUBLICA DE ANGOLA - Persona Física"
    COMBOX.AddItem "50000001500         REPUBLICA DE CABO VERDE (Estado independiente) - Persona Física"
    COMBOX.AddItem "50000001519         MOZAMBIQUE - Persona Física"
    COMBOX.AddItem "50000001527         CONGO REP.POPULAR - Persona Física"
    COMBOX.AddItem "50000001535         CHAD - Persona Física"
    COMBOX.AddItem "50000001543         MALAWI - Persona Física"
    COMBOX.AddItem "50000001551         TANZANIA - Persona Física"
    COMBOX.AddItem "50000001586         COSTA RICA - Persona Física"
    COMBOX.AddItem "50000001616         ZAIRE - Persona Física"
    COMBOX.AddItem "50000001624         BENIN - Persona Física"
    COMBOX.AddItem "50000001632         MALI - Persona Física"
    COMBOX.AddItem "50000001705         UGANDA - Persona Física"
    COMBOX.AddItem "50000001713         SUDAFRICA, REP. DE - Persona Física"
    COMBOX.AddItem "50000001810         REPUBLICA DE SEYCHELLES (Estado independiente) - Persona Física"
    COMBOX.AddItem "50000001829         SANTO TOME Y PRINCIPE - Persona Física"
    COMBOX.AddItem "50000001837         NAMIBIA - Persona Física"
    COMBOX.AddItem "50000001845         GUINEA BISSAU - Persona Física"
    COMBOX.AddItem "50000001853         ERITREA - Persona Física"
    COMBOX.AddItem "50000001861         REPUBLICA DE DJIBUTI (Estado independiente) - Persona Física"
    COMBOX.AddItem "50000001896         COMORAS - Persona Física"
    COMBOX.AddItem "50000001985         INDETERMINADO (AFRICA) - Persona Física"
    COMBOX.AddItem "50000002019         BARBADOS (Estado independiente) - Persona Física"
    COMBOX.AddItem "50000002043         CANADA - Persona Física"
    COMBOX.AddItem "50000002051         COLOMBIA - Persona Física"
    COMBOX.AddItem "50000002094         DOMINICANA, REPUBLICA - Persona Física"
    COMBOX.AddItem "50000002116         EL SALVADOR - Persona Física"
    COMBOX.AddItem "50000002124         ESTADOS UNIDOS - Persona Física"
    COMBOX.AddItem "50000002132         GUATEMALA - Persona Física"
    COMBOX.AddItem "50000002140         REPUBLICA COOPERATIVA DE GUYANA (Estado independiente) - Persona Física"
    COMBOX.AddItem "50000002159         HAITI - Persona Física"
    COMBOX.AddItem "50000002167         HONDURAS - Persona Física"
    COMBOX.AddItem "50000002175         JAMAICA - Persona Física"
    COMBOX.AddItem "50000002183         MEXICO - Persona Física"
    COMBOX.AddItem "50000002191         NICARAGUA - Persona Física"
    COMBOX.AddItem "50000002205         REPUBLICA DE PANAMA (Estado independiente) - Persona Física"
    COMBOX.AddItem "50000002213         ESTADO LIBRE ASOCIADO DE PUERTO RICO (Estado asoc. a EEUU) - Persona Física"
    COMBOX.AddItem "50000002221         PERU - Persona Física"
    COMBOX.AddItem "50000002256         ANTIGUA Y BARBUDA (Estado independiente) - Persona Física"
    COMBOX.AddItem "50000002264         VENEZUELA - Persona Física"
    COMBOX.AddItem "50000002272         POS.BRITANICA (AMERICA) - Persona Física"
    COMBOX.AddItem "50000002280         POS.DANESA (AMERICA) - Persona Física"
    COMBOX.AddItem "50000002299         POS.FRANCESA (AMERICA) - Persona Física"
    COMBOX.AddItem "50000002302         POS.PAISES BAJOS (AMERICA) - Persona Física"
    COMBOX.AddItem "50000002310         POS.E.E.U.U. (AMERICA) - Persona Física"
    COMBOX.AddItem "50000002329         SURINAME - Persona Física"
    COMBOX.AddItem "50000002337         EL COMMONWEALTH DE DOMINICA (Estado Asociado) - Persona Física"
    COMBOX.AddItem "50000002345         SANTA LUCIA - Persona Física"
    COMBOX.AddItem "50000002353         SAN VICENTE Y LAS GRANADINAS (Estado independiente) - Persona Física"
    COMBOX.AddItem "50000002361         BELICE (Estado independiente) - Persona Física"
    COMBOX.AddItem "50000002396         CUBA - Persona Física"
    COMBOX.AddItem "50000002426         ECUADOR - Persona Física"
    COMBOX.AddItem "50000002434         REPUBLICA DE TRINIDAD Y TOBAGO - Persona Física"
    COMBOX.AddItem "50000002825         BUTAN - Persona Física"
    COMBOX.AddItem "50000002841         MYANMAR (EX BIRMANIA) - Persona Física"
    COMBOX.AddItem "50000002876         ISRAEL - Persona Física"
    COMBOX.AddItem "50000002882         ESTADO ASOCIADO DE GRANADA (Estado independiente) - Persona Física"
    COMBOX.AddItem "50000002892         FEDERACION DE SAN CRISTOBAL (Islas Saint Kitts and Nevis) - Persona Física"
    COMBOX.AddItem "50000002906         COMUNIDAD DE LAS BAHAMAS (Estado independiente) - Persona Física"
    COMBOX.AddItem "50000002914         TAILANDIA - Persona Física"
    COMBOX.AddItem "50000002922         INDETERMINADO (AMERICA) - Persona Física"
    COMBOX.AddItem "50000002930         IRAN - Persona Física"
    COMBOX.AddItem "50000002981         ESTADO DE QATAR (Estado independiente) - Persona Física"
    COMBOX.AddItem "50000003007         REINO HACHEMITA DE JORDANIA - Persona Física"
    COMBOX.AddItem "50000003015         AFGANISTAN - Persona Física"
    COMBOX.AddItem "50000003023         ARABIA SAUDITA - Persona Física"
    COMBOX.AddItem "50000003031         ESTADO DE BAHREIN (Estado independiente) - Persona Física"
    COMBOX.AddItem "50000003066         CAMBOYA (EX KAMPUCHEA) - Persona Física"
    COMBOX.AddItem "50000003074         REPUBLICA DEMOCRATICA SOCIALISTA DE SRI LANKA - Persona Física"
    COMBOX.AddItem "50000003082         COREA DEMOCRATICA  - Persona Física"
    COMBOX.AddItem "50000003090         COREA REPUBLICANA - Persona Física"
    COMBOX.AddItem "50000003104         CHINA REP.POPULAR - Persona Física"
    COMBOX.AddItem "50000003112         REPUBLICA DE CHIPRE (Estado independiente) - Persona Física"
    COMBOX.AddItem "50000003120         FILIPINAS - Persona Física"
    COMBOX.AddItem "50000003139         TAIWAN - Persona Física"
    COMBOX.AddItem "50000003147         GAZA - Persona Física"
    COMBOX.AddItem "50000003155         INDIA - Persona Física"
    COMBOX.AddItem "50000003163         INDONESIA - Persona Física"
    COMBOX.AddItem "50000003171         IRAK - Persona Física"
    COMBOX.AddItem "50000003201         JAPON - Persona Física"
    COMBOX.AddItem "50000003236         ESTADO DE KUWAIT (Estado independiente) - Persona Física"
    COMBOX.AddItem "50000003244         LAOS - Persona Física"
    COMBOX.AddItem "50000003252         LIBANO - Persona Física"
    COMBOX.AddItem "50000003260         MALASIA - Persona Física"
    COMBOX.AddItem "50000003279         REPUBLICA DE MALDIVAS (Estado independiente) - Persona Física"
    COMBOX.AddItem "50000003287         SULTANATO DE OMAN - Persona Física"
    COMBOX.AddItem "50000003295         MONGOLIA - Persona Física"
    COMBOX.AddItem "50000003309         NEPAL - Persona Física"
    COMBOX.AddItem "50000003317         EMIRATOS ARABES UNIDOS (Estado independiente) - Persona Física"
    COMBOX.AddItem "50000003325         PAKISTAN - Persona Física"
    COMBOX.AddItem "50000003333         SINGAPUR - Persona Física"
    COMBOX.AddItem "50000003341         SIRIA - Persona Física"
    COMBOX.AddItem "50000003376         VIETNAM - Persona Física"
    COMBOX.AddItem "50000003392         REPUBLICA DEL YEMEN - Persona Física"
    COMBOX.AddItem "50000003414         POS.BRITANICA (HONG KONG) - Persona Física"
    COMBOX.AddItem "50000003422         POS.JAPONESA (ASIA) - Persona Física"
    COMBOX.AddItem "50000003449         MACAO - Persona Física"
    COMBOX.AddItem "50000003457         BANGLADESH - Persona Física"
    COMBOX.AddItem "50000003503         TURQUIA - Persona Física"
    COMBOX.AddItem "50000003546         ITALIA - Persona Física"
    COMBOX.AddItem "50000003554         TURKMENISTAN - Persona Física"
    COMBOX.AddItem "50000003562         UZBEKISTAN - Persona Física"
    COMBOX.AddItem "50000003570         TERRITORIOS AUTONOMOS PALESTINOS - Persona Física"
    COMBOX.AddItem "50000003813         ISLANDIA - Persona Física"
    COMBOX.AddItem "50000003880         GEORGIA - Persona Física"
    COMBOX.AddItem "50000003899         TAYIKISTAN - Persona Física"
    COMBOX.AddItem "50000003902         AZERBAIDZHAN - Persona Física"
    COMBOX.AddItem "50000003910         BRUNEI DARUSSALAM (Estado independiente) - Persona Física"
    COMBOX.AddItem "50000003929         KAZAJSTAN - Persona Física"
    COMBOX.AddItem "50000003937         KIRGUISTAN - Persona Física"
    COMBOX.AddItem "50000003961         INDETERMINADO (ASIA) - Persona Física"
    COMBOX.AddItem "50000004011         REPUBLICA DE ALBANIA - Persona Física"
    COMBOX.AddItem "50000004046         PRINCIPADO DEL VALLE DE ANDORRA - Persona Física"
    COMBOX.AddItem "50000004054         AUSTRIA - Persona Física"
    COMBOX.AddItem "50000004062         BELGICA - Persona Física"
    COMBOX.AddItem "50000004070         BULGARIA - Persona Física"
    COMBOX.AddItem "50000004097         DINAMARCA - Persona Física"
    COMBOX.AddItem "50000004100         ESPAÑA - Persona Física"
    COMBOX.AddItem "50000004119         FINLANDIA - Persona Física"
    COMBOX.AddItem "50000004127         FRANCIA - Persona Física"
    COMBOX.AddItem "50000004135         GRECIA - Persona Física"
    COMBOX.AddItem "50000004143         HUNGRIA - Persona Física"
    COMBOX.AddItem "50000004151         IRLANDA (EIRE) - Persona Física"
    COMBOX.AddItem "50000004186         PRINCIPADO DE LIECHTENSTEIN (Estado independiente) - Persona Física"
    COMBOX.AddItem "50000004194         GRAN DUCADO DE LUXEMBURGO - Persona Física"
    COMBOX.AddItem "50000004216         PRINCIPADO DE MONACO - Persona Física"
    COMBOX.AddItem "50000004224         NORUEGA - Persona Física"
    COMBOX.AddItem "50000004232         PAISES BAJOS - Persona Física"
    COMBOX.AddItem "50000004240         POLONIA - Persona Física"
    COMBOX.AddItem "50000004259         PORTUGAL - Persona Física"
    COMBOX.AddItem "50000004267         REINO UNIDO - Persona Física"
    COMBOX.AddItem "50000004275         RUMANIA - Persona Física"
    COMBOX.AddItem "50000004283         SERENISIMA REPUBLICA DE SAN MARINO (Estado independiente) - Persona Física"
    COMBOX.AddItem "50000004291         SUECIA - Persona Física"
    COMBOX.AddItem "50000004305         SUIZA - Persona Física"
    COMBOX.AddItem "50000004313         SANTA SEDE (VATICANO) - Persona Física"
    COMBOX.AddItem "50000004321         YUGOSLAVIA - Persona Física"
    COMBOX.AddItem "50000004364         REPUBLICA DE MALTA (Estado independiente) - Persona Física"
    COMBOX.AddItem "50000004380         ALEMANIA, REP. FED. - Persona Física"
    COMBOX.AddItem "50000004399         BIELORUSIA - Persona Física"
    COMBOX.AddItem "50000004402         ESTONIA - Persona Física"
    COMBOX.AddItem "50000004410         LETONIA - Persona Física"
    COMBOX.AddItem "50000004429         LITUANIA - Persona Física"
    COMBOX.AddItem "50000004437         MOLDOVA - Persona Física"
    COMBOX.AddItem "50000004461         BOSNIA HERZEGOVINA - Persona Física"
    COMBOX.AddItem "50000004496         ESLOVENIA - Persona Física"
    COMBOX.AddItem "50000004909         MACEDONIA - Persona Física"
    COMBOX.AddItem "50000004917         POS.BRITANICA (EUROPA) - Persona Física"
    COMBOX.AddItem "50000004984         INDETERMINADO (EUROPA) - Persona Física"
    COMBOX.AddItem "50000004992         AUSTRALIA - Persona Física"
    COMBOX.AddItem "50000005034         REPUBLICA DE NAURU (Estado independiente) - Persona Física"
    COMBOX.AddItem "50000005042         NUEVA ZELANDA - Persona Física"
    COMBOX.AddItem "50000005050         REPUBLICA DE VANUATU - Persona Física"
    COMBOX.AddItem "50000005069         SAMOA OCCIDENTAL - Persona Física"
    COMBOX.AddItem "50000005077         POS.AUSTRALIANA (OCEANIA) - Persona Física"
    COMBOX.AddItem "50000005085         POS.BRITANICA (OCEANIA) - Persona Física"
    COMBOX.AddItem "50000005093         POS.FRANCESA (OCEANIA) - Persona Física"
    COMBOX.AddItem "50000005107         POS.NEOCELANDESA (OCEANIA) - Persona Física"
    COMBOX.AddItem "50000005115         POS.E.E.U.U. (OCEANIA) - Persona Física"
    COMBOX.AddItem "50000005123         FIJI, ISLAS - Persona Física"
    COMBOX.AddItem "50000005131         PAPUA, ISLAS - Persona Física"
    COMBOX.AddItem "50000005166         KIRIBATI - Persona Física"
    COMBOX.AddItem "50000005174         TUVALU - Persona Física"
    COMBOX.AddItem "50000005182         ISLAS SALOMON - Persona Física"
    COMBOX.AddItem "50000005190         REINO DE TONGA (Estado independiente) - Persona Física"
    COMBOX.AddItem "50000005204         REPUBLICA DE LAS ISLAS MARSHALL (Estado independiente) - Persona Física"
    COMBOX.AddItem "50000005212         ISLAS MARIANAS - Persona Física"
    COMBOX.AddItem "50000005905         MICRONESIA ESTADOS FED. - Persona Física"
    COMBOX.AddItem "50000005913         PALAU - Persona Física"
    COMBOX.AddItem "50000005980         INDETERMINADO (OCEANIA) - Persona Física"
    COMBOX.AddItem "50000006014         RUSA, FEDERACION - Persona Física"
    COMBOX.AddItem "50000006022         ARMENIA - Persona Física"
    COMBOX.AddItem "50000006030         CROACIA - Persona Física"
    COMBOX.AddItem "50000006049         UCRANIA - Persona Física"
    COMBOX.AddItem "50000006057         CHECA, REPUBLICA - Persona Física"
    COMBOX.AddItem "50000006065         ESLOVACA, REPUBLICA - Persona Física"
    COMBOX.AddItem "50000006529         ANGUILA (Territorio no autónomo del Reino Unido) - Persona Física"
    COMBOX.AddItem "50000006537         ARUBA (Territorio de Países Bajos) - Persona Física"
    COMBOX.AddItem "50000006545         ISLAS DE COOK (Territorio autónomo asociado a Nueva Zelanda) - Persona Física"
    COMBOX.AddItem "50000006553         PATAU - Persona Física"
    COMBOX.AddItem "50000006561         POLINESIA FRANCESA (Territorio de Ultramar de Francia) - Persona Física"
    COMBOX.AddItem "50000006596         ANTILLAS HOLANDESAS (Territorio de Países Bajos) - Persona Física"
    COMBOX.AddItem "50000006626         ASCENCION - Persona Física"
    COMBOX.AddItem "50000006634         BERMUDAS (Territorio no autónomo del Reino Unido) - Persona Física"
    COMBOX.AddItem "50000006642         CAMPIONE D@ITALIA - Persona Física"
    COMBOX.AddItem "50000006650         COLONIA DE GIBRALTAR - Persona Física"
    COMBOX.AddItem "50000006669         GROENLANDIA - Persona Física"
    COMBOX.AddItem "50000006677         GUAM (Territorio no autónomo de los EEUU) - Persona Física"
    COMBOX.AddItem "50000006685         HONK KONG (Territorio de China) - Persona Física"
    COMBOX.AddItem "50000006693         ISLAS AZORES - Persona Física"
    COMBOX.AddItem "50000006707         ISLAS DEL CANAL:Guernesey,Jersey,Alderney,G.Stark,L.Sark,etc - Persona Física"
    COMBOX.AddItem "50000006715         ISLAS CAIMAN (Territorio no autónomo del Reino Unido) - Persona Física"
    COMBOX.AddItem "50000006723         ISLA CHRISTMAS - Persona Física"
    COMBOX.AddItem "50000006731         ISLA DE COCOS O KEELING - Persona Física"
    COMBOX.AddItem "50000006766         ISLA DE MAN (Territorio del Reino Unido) - Persona Física"
    COMBOX.AddItem "50000006774         ISLA DE NORFOLK - Persona Física"
    COMBOX.AddItem "50000006782         ISLAS TURKAS Y CAICOS (Territorio no autónomo del R. Unido) - Persona Física"
    COMBOX.AddItem "50000006790         ISLAS PACIFICO - Persona Física"
    COMBOX.AddItem "50000006804         ISLA DE SAN PEDRO Y MIGUELON - Persona Física"
    COMBOX.AddItem "50000006812         ISLA QESHM - Persona Física"
    COMBOX.AddItem "50000006820         ISLAS VIRGENES BRITANICAS(Territorio no autónomo de R.UNIDO) - Persona Física"
    COMBOX.AddItem "50000006839         ISLAS VIRGENES DE ESTADOS UNIDOS DE AMERICA - Persona Física"
    COMBOX.AddItem "50000006847         LABUAN - Persona Física"
    COMBOX.AddItem "50000006855         MADEIRA (Territorio de Portugal) - Persona Física"
    COMBOX.AddItem "50000006863         MONTSERRAT (Territorio no autónomo del Reino Unido) - Persona Física"
    COMBOX.AddItem "50000006871         NIUE - Persona Física"
    COMBOX.AddItem "50000006901         PITCAIRN - Persona Física"
    COMBOX.AddItem "50000006936         REGIMEN APLICABLE A LAS SA FINANCIERAS(ley 11.073 de la ROU) - Persona Física"
    COMBOX.AddItem "50000006944         SANTA ELENA - Persona Física"
    COMBOX.AddItem "50000006952         SAMOA AMERICANA (Territorio no autónomo de los EEUU) - Persona Física"
    COMBOX.AddItem "50000006960         ARCHIPIELAGO DE SVBALBARD - Persona Física"
    COMBOX.AddItem "50000006979         TRISTAN DA CUNHA - Persona Física"
    COMBOX.AddItem "50000006987         TRIESTE (Italia) - Persona Física"
    COMBOX.AddItem "50000006995         TOKELAU - Persona Física"
    COMBOX.AddItem "50000007002         ZONA LIBRE DE OSTRAVA (ciudad de la antigua Checoeslovaquia) - Persona Física"
    COMBOX.AddItem "50000009986         PARA PERSONAS FISICAS DE INDETERMINADO (CONTINENTE) - Persona Física"
    COMBOX.AddItem "50000009994         PARA PERSONAS FISICAS DE OTROS PAISES - Persona Física"
    COMBOX.AddItem "51600000016         URUGUAY - Otro tipo de Entidad"
    COMBOX.AddItem "51600000024         PARAGUAY - Otro tipo de Entidad"
    COMBOX.AddItem "51600000032         CHILE - Otro tipo de Entidad"
    COMBOX.AddItem "51600000040         BOLIVIA - Otro tipo de Entidad"
    COMBOX.AddItem "51600000059         BRASIL - Otro tipo de Entidad"
    COMBOX.AddItem "51600001012         BURKINA FASO - Otro tipo de Entidad"
    COMBOX.AddItem "51600001020         ARGELIA - Otro tipo de Entidad"
    COMBOX.AddItem "51600001039         BOTSWANA - Otro tipo de Entidad"
    COMBOX.AddItem "51600001047         BURUNDI - Otro tipo de Entidad"
    COMBOX.AddItem "51600001055         CAMERUN - Otro tipo de Entidad"
    COMBOX.AddItem "51600001071         CENTRO AFRICANO, REP. - Otro tipo de Entidad"
    COMBOX.AddItem "51600001101         COSTA DE MARFIL - Otro tipo de Entidad"
    COMBOX.AddItem "51600001136         EGIPTO - Otro tipo de Entidad"
    COMBOX.AddItem "51600001144         ETIOPIA - Otro tipo de Entidad"
    COMBOX.AddItem "51600001152         GABON - Otro tipo de Entidad"
    COMBOX.AddItem "51600001160         GAMBIA - Otro tipo de Entidad"
    COMBOX.AddItem "51600001179         GHANA - Otro tipo de Entidad"
    COMBOX.AddItem "51600001187         GUINEA - Otro tipo de Entidad"
    COMBOX.AddItem "51600001195         GUINEA ECUATORIAL - Otro tipo de Entidad"
    COMBOX.AddItem "51600001209         KENIA - Otro tipo de Entidad"
    COMBOX.AddItem "51600001217         LESOTHO - Otro tipo de Entidad"
    COMBOX.AddItem "51600001225         REPUBLICA DE LIBERIA (Estado independiente) - Otro tipo de Entidad"
    COMBOX.AddItem "51600001233         LIBIA - Otro tipo de Entidad"
    COMBOX.AddItem "51600001241         MADAGASCAR - Otro tipo de Entidad"
    COMBOX.AddItem "51600001276         MARRUECOS - Otro tipo de Entidad"
    COMBOX.AddItem "51600001284         REPUBLICA DE MAURICIO - Otro tipo de Entidad"
    COMBOX.AddItem "51600001292         MAURITANIA - Otro tipo de Entidad"
    COMBOX.AddItem "51600001306         NIGER - Otro tipo de Entidad"
    COMBOX.AddItem "51600001314         NIGERIA - Otro tipo de Entidad"
    COMBOX.AddItem "51600001322         ZIMBABWE - Otro tipo de Entidad"
    COMBOX.AddItem "51600001330         RUANDA - Otro tipo de Entidad"
    COMBOX.AddItem "51600001349         SENEGAL - Otro tipo de Entidad"
    COMBOX.AddItem "51600001357         SIERRA LEONA - Otro tipo de Entidad"
    COMBOX.AddItem "51600001365         SOMALIA - Otro tipo de Entidad"
    COMBOX.AddItem "51600001373         REINO DE SWAZILANDIA (Estado independiente) - Otro tipo de Entidad"
    COMBOX.AddItem "51600001381         SUDAN - Otro tipo de Entidad"
    COMBOX.AddItem "51600001403         TOGO - Otro tipo de Entidad"
    COMBOX.AddItem "51600001411         REPUBLICA TUNECINA - Otro tipo de Entidad"
    COMBOX.AddItem "51600001446         ZAMBIA - Otro tipo de Entidad"
    COMBOX.AddItem "51600001454         POS.BRITANICA (AFRICA) - Otro tipo de Entidad"
    COMBOX.AddItem "51600001462         POS.ESPAÑOLA (AFRICA) - Otro tipo de Entidad"
    COMBOX.AddItem "51600001470         POS.FRANCESA (AFRICA) - Otro tipo de Entidad"
    COMBOX.AddItem "51600001489         POS.PORTUGUESA (AFRICA) - Otro tipo de Entidad"
    COMBOX.AddItem "51600001497         REPUBLICA DE ANGOLA - Otro tipo de Entidad"
    COMBOX.AddItem "51600001500         REPUBLICA DE CABO VERDE (Estado independiente) - Otro tipo de Entidad"
    COMBOX.AddItem "51600001519         MOZAMBIQUE - Otro tipo de Entidad"
    COMBOX.AddItem "51600001527         CONGO REP.POPULAR - Otro tipo de Entidad"
    COMBOX.AddItem "51600001535         CHAD - Otro tipo de Entidad"
    COMBOX.AddItem "51600001543         MALAWI - Otro tipo de Entidad"
    COMBOX.AddItem "51600001551         TANZANIA - Otro tipo de Entidad"
    COMBOX.AddItem "51600001586         COSTA RICA - Otro tipo de Entidad"
    COMBOX.AddItem "51600001616         ZAIRE - Otro tipo de Entidad"
    COMBOX.AddItem "51600001624         BENIN - Otro tipo de Entidad"
    COMBOX.AddItem "51600001632         MALI - Otro tipo de Entidad"
    COMBOX.AddItem "51600001705         UGANDA - Otro tipo de Entidad"
    COMBOX.AddItem "51600001713         SUDAFRICA, REP. DE - Otro tipo de Entidad"
    COMBOX.AddItem "51600001810         REPUBLICA DE SEYCHELLES (Estado independiente) - Otro tipo de Entidad"
    COMBOX.AddItem "51600001829         SANTO TOME Y PRINCIPE - Otro tipo de Entidad"
    COMBOX.AddItem "51600001837         NAMIBIA - Otro tipo de Entidad"
    COMBOX.AddItem "51600001845         GUINEA BISSAU - Otro tipo de Entidad"
    COMBOX.AddItem "51600001853         ERITREA - Otro tipo de Entidad"
    COMBOX.AddItem "51600001861         REPUBLICA DE DJIBUTI (Estado independiente) - Otro tipo de Entidad"
    COMBOX.AddItem "51600001896         COMORAS - Otro tipo de Entidad"
    COMBOX.AddItem "51600001985         INDETERMINADO (AFRICA) - Otro tipo de Entidad"
    COMBOX.AddItem "51600002019         BARBADOS (Estado independiente) - Otro tipo de Entidad"
    COMBOX.AddItem "51600002043         CANADA - Otro tipo de Entidad"
    COMBOX.AddItem "51600002051         COLOMBIA - Otro tipo de Entidad"
    COMBOX.AddItem "51600002094         DOMINICANA, REPUBLICA - Otro tipo de Entidad"
    COMBOX.AddItem "51600002116         EL SALVADOR - Otro tipo de Entidad"
    COMBOX.AddItem "51600002124         ESTADOS UNIDOS - Otro tipo de Entidad"
    COMBOX.AddItem "51600002132         GUATEMALA - Otro tipo de Entidad"
    COMBOX.AddItem "51600002140         REPUBLICA COOPERATIVA DE GUYANA (Estado independiente) - Otro tipo de Entidad"
    COMBOX.AddItem "51600002159         HAITI - Otro tipo de Entidad"
    COMBOX.AddItem "51600002167         HONDURAS - Otro tipo de Entidad"
    COMBOX.AddItem "51600002175         JAMAICA - Otro tipo de Entidad"
    COMBOX.AddItem "51600002183         MEXICO - Otro tipo de Entidad"
    COMBOX.AddItem "51600002191         NICARAGUA - Otro tipo de Entidad"
    COMBOX.AddItem "51600002205         REPUBLICA DE PANAMA (Estado independiente) - Otro tipo de Entidad"
    COMBOX.AddItem "51600002213         ESTADO LIBRE ASOCIADO DE PUERTO RICO (Estado asoc. a EEUU) - Otro tipo de Entidad"
    COMBOX.AddItem "51600002221         PERU - Otro tipo de Entidad"
    COMBOX.AddItem "51600002256         ANTIGUA Y BARBUDA (Estado independiente) - Otro tipo de Entidad"
    COMBOX.AddItem "51600002264         VENEZUELA - Otro tipo de Entidad"
    COMBOX.AddItem "51600002272         POS.BRITANICA (AMERICA) - Otro tipo de Entidad"
    COMBOX.AddItem "51600002280         POS.DANESA (AMERICA) - Otro tipo de Entidad"
    COMBOX.AddItem "51600002299         POS.FRANCESA (AMERICA) - Otro tipo de Entidad"
    COMBOX.AddItem "51600002302         POS.PAISES BAJOS (AMERICA) - Otro tipo de Entidad"
    COMBOX.AddItem "51600002310         POS.E.E.U.U. (AMERICA) - Otro tipo de Entidad"
    COMBOX.AddItem "51600002329         SURINAME - Otro tipo de Entidad"
    COMBOX.AddItem "51600002337         EL COMMONWEALTH DE DOMINICA (Estado Asociado) - Otro tipo de Entidad"
    COMBOX.AddItem "51600002345         SANTA LUCIA - Otro tipo de Entidad"
    COMBOX.AddItem "51600002353         SAN VICENTE Y LAS GRANADINAS (Estado independiente) - Otro tipo de Entidad"
    COMBOX.AddItem "51600002361         BELICE (Estado independiente) - Otro tipo de Entidad"
    COMBOX.AddItem "51600002396         CUBA - Otro tipo de Entidad"
    COMBOX.AddItem "51600002426         ECUADOR - Otro tipo de Entidad"
    COMBOX.AddItem "51600002434         REPUBLICA DE TRINIDAD Y TOBAGO - Otro tipo de Entidad"
    COMBOX.AddItem "51600002825         BUTAN - Otro tipo de Entidad"
    COMBOX.AddItem "51600002841         MYANMAR (EX BIRMANIA) - Otro tipo de Entidad"
    COMBOX.AddItem "51600002876         ISRAEL - Otro tipo de Entidad"
    COMBOX.AddItem "51600002884         ESTADO ASOCIADO DE GRANADA (Estado independiente) - Otro tipo de Entidad"
    COMBOX.AddItem "51600002892         FEDERACION DE SAN CRISTOBAL (Islas Saint Kitts and Nevis) - Otro tipo de Entidad"
    COMBOX.AddItem "51600002906         COMUNIDAD DE LAS BAHAMAS (Estado independiente) - Otro tipo de Entidad"
    COMBOX.AddItem "51600002914         TAILANDIA - Otro tipo de Entidad"
    COMBOX.AddItem "51600002922         INDETERMINADO (AMERICA) - Otro tipo de Entidad"
    COMBOX.AddItem "51600002930         IRAN - Otro tipo de Entidad"
    COMBOX.AddItem "51600002981         ESTADO DE QATAR (Estado independiente) - Otro tipo de Entidad"
    COMBOX.AddItem "51600003007         REINO HACHEMITA DE JORDANIA - Otro tipo de Entidad"
    COMBOX.AddItem "51600003015         AFGANISTAN - Otro tipo de Entidad"
    COMBOX.AddItem "51600003023         ARABIA SAUDITA - Otro tipo de Entidad"
    COMBOX.AddItem "51600003031         ESTADO DE BAHREIN (Estado independiente) - Otro tipo de Entidad"
    COMBOX.AddItem "51600003066         CAMBOYA (EX KAMPUCHEA) - Otro tipo de Entidad"
    COMBOX.AddItem "51600003074         REPUBLICA DEMOCRATICA SOCIALISTA DE SRI LANKA - Otro tipo de Entidad"
    COMBOX.AddItem "51600003082         COREA DEMOCRATICA  - Otro tipo de Entidad"
    COMBOX.AddItem "51600003090         COREA REPUBLICANA - Otro tipo de Entidad"
    COMBOX.AddItem "51600003104         CHINA REP.POPULAR - Otro tipo de Entidad"
    COMBOX.AddItem "51600003112         REPUBLICA DE CHIPRE (Estado independiente) - Otro tipo de Entidad"
    COMBOX.AddItem "51600003120         FILIPINAS - Otro tipo de Entidad"
    COMBOX.AddItem "51600003139         TAIWAN - Otro tipo de Entidad"
    COMBOX.AddItem "51600003147         GAZA - Otro tipo de Entidad"
    COMBOX.AddItem "51600003155         INDIA - Otro tipo de Entidad"
    COMBOX.AddItem "51600003163         INDONESIA - Otro tipo de Entidad"
    COMBOX.AddItem "51600003171         IRAK - Otro tipo de Entidad"
    COMBOX.AddItem "51600003201         JAPON - Otro tipo de Entidad"
    COMBOX.AddItem "51600003236         ESTADO DE KUWAIT (Estado independiente) - Otro tipo de Entidad"
    COMBOX.AddItem "51600003244         LAOS - Otro tipo de Entidad"
    COMBOX.AddItem "51600003252         LIBANO - Otro tipo de Entidad"
    COMBOX.AddItem "51600003260         MALASIA - Otro tipo de Entidad"
    COMBOX.AddItem "51600003279         REPUBLICA DE MALDIVAS (Estado independiente) - Otro tipo de Entidad"
    COMBOX.AddItem "51600003287         SULTANATO DE OMAN - Otro tipo de Entidad"
    COMBOX.AddItem "51600003295         MONGOLIA - Otro tipo de Entidad"
    COMBOX.AddItem "51600003309         NEPAL - Otro tipo de Entidad"
    COMBOX.AddItem "51600003317         EMIRATOS ARABES UNIDOS (Estado independiente) - Otro tipo de Entidad"
    COMBOX.AddItem "51600003325         PAKISTAN - Otro tipo de Entidad"
    COMBOX.AddItem "51600003333         SINGAPUR - Otro tipo de Entidad"
    COMBOX.AddItem "51600003341         SIRIA - Otro tipo de Entidad"
    COMBOX.AddItem "51600003376         VIETNAM - Otro tipo de Entidad"
    COMBOX.AddItem "51600003392         REPUBLICA DEL YEMEN - Otro tipo de Entidad"
    COMBOX.AddItem "51600003414         POS.BRITANICA (HONG KONG) - Otro tipo de Entidad"
    COMBOX.AddItem "51600003422         POS.JAPONESA (ASIA) - Otro tipo de Entidad"
    COMBOX.AddItem "51600003449         MACAO - Otro tipo de Entidad"
    COMBOX.AddItem "51600003457         BANGLADESH - Otro tipo de Entidad"
    COMBOX.AddItem "51600003503         TURQUIA - Otro tipo de Entidad"
    COMBOX.AddItem "51600003546         ITALIA - Otro tipo de Entidad"
    COMBOX.AddItem "51600003554         TURKMENISTAN - Otro tipo de Entidad"
    COMBOX.AddItem "51600003562         UZBEKISTAN - Otro tipo de Entidad"
    COMBOX.AddItem "51600003570         TERRITORIOS AUTONOMOS PALESTINOS - Otro tipo de Entidad"
    COMBOX.AddItem "51600003813         ISLANDIA - Otro tipo de Entidad"
    COMBOX.AddItem "51600003880         GEORGIA - Otro tipo de Entidad"
    COMBOX.AddItem "51600003899         TAYIKISTAN - Otro tipo de Entidad"
    COMBOX.AddItem "51600003902         AZERBAIDZHAN - Otro tipo de Entidad"
    COMBOX.AddItem "51600003910         BRUNEI DARUSSALAM (Estado independiente) - Otro tipo de Entidad"
    COMBOX.AddItem "51600003929         KAZAJSTAN - Otro tipo de Entidad"
    COMBOX.AddItem "51600003937         KIRGUISTAN - Otro tipo de Entidad"
    COMBOX.AddItem "51600003961         INDETERMINADO (ASIA) - Otro tipo de Entidad"
    COMBOX.AddItem "51600004011         REPUBLICA DE ALBANIA - Otro tipo de Entidad"
    COMBOX.AddItem "51600004046         PRINCIPADO DEL VALLE DE ANDORRA - Otro tipo de Entidad"
    COMBOX.AddItem "51600004054         AUSTRIA - Otro tipo de Entidad"
    COMBOX.AddItem "51600004062         BELGICA - Otro tipo de Entidad"
    COMBOX.AddItem "51600004070         BULGARIA - Otro tipo de Entidad"
    COMBOX.AddItem "51600004097         DINAMARCA - Otro tipo de Entidad"
    COMBOX.AddItem "51600004100         ESPAÑA - Otro tipo de Entidad"
    COMBOX.AddItem "51600004119         FINLANDIA - Otro tipo de Entidad"
    COMBOX.AddItem "51600004127         FRANCIA - Otro tipo de Entidad"
    COMBOX.AddItem "51600004135         GRECIA - Otro tipo de Entidad"
    COMBOX.AddItem "51600004143         HUNGRIA - Otro tipo de Entidad"
    COMBOX.AddItem "51600004151         IRLANDA (EIRE) - Otro tipo de Entidad"
    COMBOX.AddItem "51600004186         PRINCIPADO DE LIECHTENSTEIN (Estado independiente) - Otro tipo de Entidad"
    COMBOX.AddItem "51600004194         GRAN DUCADO DE LUXEMBURGO - Otro tipo de Entidad"
    COMBOX.AddItem "51600004216         PRINCIPADO DE MONACO - Otro tipo de Entidad"
    COMBOX.AddItem "51600004224         NORUEGA - Otro tipo de Entidad"
    COMBOX.AddItem "51600004232         PAISES BAJOS - Otro tipo de Entidad"
    COMBOX.AddItem "51600004240         POLONIA - Otro tipo de Entidad"
    COMBOX.AddItem "51600004259         PORTUGAL - Otro tipo de Entidad"
    COMBOX.AddItem "51600004267         REINO UNIDO - Otro tipo de Entidad"
    COMBOX.AddItem "51600004275         RUMANIA - Otro tipo de Entidad"
    COMBOX.AddItem "51600004283         SERENISIMA REPUBLICA DE SAN MARINO (Estado independiente) - Otro tipo de Entidad"
    COMBOX.AddItem "51600004291         SUECIA - Otro tipo de Entidad"
    COMBOX.AddItem "51600004305         SUIZA - Otro tipo de Entidad"
    COMBOX.AddItem "51600004313         SANTA SEDE (VATICANO) - Otro tipo de Entidad"
    COMBOX.AddItem "51600004321         YUGOSLAVIA - Otro tipo de Entidad"
    COMBOX.AddItem "51600004364         REPUBLICA DE MALTA (Estado independiente) - Otro tipo de Entidad"
    COMBOX.AddItem "51600004380         ALEMANIA, REP. FED. - Otro tipo de Entidad"
    COMBOX.AddItem "51600004399         BIELORUSIA - Otro tipo de Entidad"
    COMBOX.AddItem "51600004402         ESTONIA - Otro tipo de Entidad"
    COMBOX.AddItem "51600004410         LETONIA - Otro tipo de Entidad"
    COMBOX.AddItem "51600004429         LITUANIA - Otro tipo de Entidad"
    COMBOX.AddItem "51600004437         MOLDOVA - Otro tipo de Entidad"
    COMBOX.AddItem "51600004461         BOSNIA HERZEGOVINA - Otro tipo de Entidad"
    COMBOX.AddItem "51600004496         ESLOVENIA - Otro tipo de Entidad"
    COMBOX.AddItem "51600004909         MACEDONIA - Otro tipo de Entidad"
    COMBOX.AddItem "51600004917         POS.BRITANICA (EUROPA) - Otro tipo de Entidad"
    COMBOX.AddItem "51600004984         INDETERMINADO (EUROPA) - Otro tipo de Entidad"
    COMBOX.AddItem "51600004992         AUSTRALIA - Otro tipo de Entidad"
    COMBOX.AddItem "51600005034         REPUBLICA DE NAURU (Estado independiente) - Otro tipo de Entidad"
    COMBOX.AddItem "51600005042         NUEVA ZELANDA - Otro tipo de Entidad"
    COMBOX.AddItem "51600005050         REPUBLICA DE VANUATU - Otro tipo de Entidad"
    COMBOX.AddItem "51600005069         SAMOA OCCIDENTAL - Otro tipo de Entidad"
    COMBOX.AddItem "51600005077         POS.AUSTRALIANA (OCEANIA) - Otro tipo de Entidad"
    COMBOX.AddItem "51600005085         POS.BRITANICA (OCEANIA) - Otro tipo de Entidad"
    COMBOX.AddItem "51600005093         POS.FRANCESA (OCEANIA) - Otro tipo de Entidad"
    COMBOX.AddItem "51600005107         POS.NEOCELANDESA (OCEANIA) - Otro tipo de Entidad"
    COMBOX.AddItem "51600005115         POS.E.E.U.U. (OCEANIA) - Otro tipo de Entidad"
    COMBOX.AddItem "51600005123         FIJI, ISLAS - Otro tipo de Entidad"
    COMBOX.AddItem "51600005131         PAPUA, ISLAS - Otro tipo de Entidad"
    COMBOX.AddItem "51600005166         KIRIBATI - Otro tipo de Entidad"
    COMBOX.AddItem "51600005174         TUVALU - Otro tipo de Entidad"
    COMBOX.AddItem "51600005182         ISLAS SALOMON - Otro tipo de Entidad"
    COMBOX.AddItem "51600005190         REINO DE TONGA (Estado independiente) - Otro tipo de Entidad"
    COMBOX.AddItem "51600005204         REPUBLICA DE LAS ISLAS MARSHALL (Estado independiente) - Otro tipo de Entidad"
    COMBOX.AddItem "51600005212         ISLAS MARIANAS - Otro tipo de Entidad"
    COMBOX.AddItem "51600005905         MICRONESIA ESTADOS FEDERADOS - Otro tipo de Entidad"
    COMBOX.AddItem "51600005913         PALAU - Otro tipo de Entidad"
    COMBOX.AddItem "51600005980         INDETERMINADO (OCEANIA) - Otro tipo de Entidad"
    COMBOX.AddItem "51600006014         RUSA, FEDERACION - Otro tipo de Entidad"
    COMBOX.AddItem "51600006022         ARMENIA - Otro tipo de Entidad"
    COMBOX.AddItem "51600006030         CROACIA - Otro tipo de Entidad"
    COMBOX.AddItem "51600006049         UCRANIA - Otro tipo de Entidad"
    COMBOX.AddItem "51600006057         CHECA, REPUBLICA - Otro tipo de Entidad"
    COMBOX.AddItem "51600006065         ESLOVACA, REPUBLICA - Otro tipo de Entidad"
    COMBOX.AddItem "51600006529         ANGUILA (Territorio no autónomo del Reino Unido) - Otro tipo de Entidad"
    COMBOX.AddItem "51600006537         ARUBA (Territorio de Países Bajos) - Otro tipo de Entidad"
    COMBOX.AddItem "51600006545         ISLAS DE COOK (Territorio autónomo asociado a Nueva Zelanda) - Otro tipo de Entidad"
    COMBOX.AddItem "51600006553         PATAU - Otro tipo de Entidad"
    COMBOX.AddItem "51600006561         POLINESIA FRANCESA (Territorio de Ultramar de Francia) - Otro tipo de Entidad"
    COMBOX.AddItem "51600006596         ANTILLAS HOLANDESAS (Territorio de Países Bajos) - Otro tipo de Entidad"
    COMBOX.AddItem "51600006626         ASCENCION - Otro tipo de Entidad"
    COMBOX.AddItem "51600006634         BERMUDAS (Territorio no autónomo del Reino Unido) - Otro tipo de Entidad"
    COMBOX.AddItem "51600006642         CAMPIONE D@ITALIA - Otro tipo de Entidad"
    COMBOX.AddItem "51600006650         COLONIA DE GIBRALTAR - Otro tipo de Entidad"
    COMBOX.AddItem "51600006669         GROENLANDIA - Otro tipo de Entidad"
    COMBOX.AddItem "51600006677         GUAM (Territorio no autónomo de los EEUU) - Otro tipo de Entidad"
    COMBOX.AddItem "51600006685         HONK KONG (Territorio de China) - Otro tipo de Entidad"
    COMBOX.AddItem "51600006693         ISLAS AZORES - Otro tipo de Entidad"
    COMBOX.AddItem "51600006707         ISLAS DEL CANAL:Guernesey,Jersey,Alderney,G.Stark,L.Sark,etc - Otro tipo de Entidad"
    COMBOX.AddItem "51600006715         ISLAS CAIMAN (Territorio no autónomo del Reino Unido) - Otro tipo de Entidad"
    COMBOX.AddItem "51600006723         ISLA CHRISTMAS - Otro tipo de Entidad"
    COMBOX.AddItem "51600006731         ISLA DE COCOS O KEELING - Otro tipo de Entidad"
    COMBOX.AddItem "51600006766         ISLA DE MAN (Territorio del Reino Unido) - Otro tipo de Entidad"
    COMBOX.AddItem "51600006774         ISLA DE NORFOLK - Otro tipo de Entidad"
    COMBOX.AddItem "51600006782         ISLAS TURKAS Y CAICOS (Territorio no autónomo del R. Unido) - Otro tipo de Entidad"
    COMBOX.AddItem "51600006790         ISLAS PACIFICO - Otro tipo de Entidad"
    COMBOX.AddItem "51600006804         ISLA DE SAN PEDRO Y MIGUELON - Otro tipo de Entidad"
    COMBOX.AddItem "51600006812         ISLA QESHM - Otro tipo de Entidad"
    COMBOX.AddItem "51600006820         ISLAS VIRGENES BRITANICAS(Territorio no autónomo de R.UNIDO) - Otro tipo de Entidad"
    COMBOX.AddItem "51600006839         ISLAS VIRGENES DE ESTADOS UNIDOS DE AMERICA - Otro tipo de Entidad"
    COMBOX.AddItem "51600006847         LABUAN - Otro tipo de Entidad"
    COMBOX.AddItem "51600006855         MADEIRA (Territorio de Portugal) - Otro tipo de Entidad"
    COMBOX.AddItem "51600006863         MONTSERRAT (Territorio no autónomo del Reino Unido) - Otro tipo de Entidad"
    COMBOX.AddItem "51600006871         NIUE - Otro tipo de Entidad"
    COMBOX.AddItem "51600006901         PITCAIRN - Otro tipo de Entidad"
    COMBOX.AddItem "51600006936         REGIMEN APLICABLE A LAS SA FINANCIERAS(ley 11.073 de la ROU) - Otro tipo de Entidad"
    COMBOX.AddItem "51600006944         SANTA ELENA - Otro tipo de Entidad"
    COMBOX.AddItem "51600006952         SAMOA AMERICANA (Territorio no autónomo de los EEUU) - Otro tipo de Entidad"
    COMBOX.AddItem "51600006960         ARCHIPIELAGO DE SVBALBARD - Otro tipo de Entidad"
    COMBOX.AddItem "51600006979         TRISTAN DA CUNHA - Otro tipo de Entidad"
    COMBOX.AddItem "51600006987         TRIESTE (Italia) - Otro tipo de Entidad"
    COMBOX.AddItem "51600006995         TOKELAU - Otro tipo de Entidad"
    COMBOX.AddItem "51600007002         ZONA LIBRE DE OSTRAVA (ciudad de la antigua Checoeslovaquia) - Otro tipo de Entidad"
    COMBOX.AddItem "51600009986         PARA PERSONAS FISICAS DE INDETERMINADO (CONTINENTE) - Otro tipo de Entidad"
    COMBOX.AddItem "51600009994         PARA PERSONAS FISICAS DE OTROS PAISES - Otro tipo de Entidad"
    COMBOX.AddItem "55000000018         URUGUAY - Persona Jurídica"
    COMBOX.AddItem "55000000026         PARAGUAY - Persona Jurídica"
    COMBOX.AddItem "55000000034         CHILE - Persona Jurídica"
    COMBOX.AddItem "55000000042         BOLIVIA - Persona Jurídica"
    COMBOX.AddItem "55000000050         BRASIL - Persona Jurídica"
    COMBOX.AddItem "55000001014         BURKINA FASO - Persona Jurídica"
    COMBOX.AddItem "55000001022         ARGELIA - Persona Jurídica"
    COMBOX.AddItem "55000001030         BOTSWANA - Persona Jurídica"
    COMBOX.AddItem "55000001049         BURUNDI - Persona Jurídica"
    COMBOX.AddItem "55000001057         CAMERUN - Persona Jurídica"
    COMBOX.AddItem "55000001073         CENTRO AFRICANO, REP. - Persona Jurídica"
    COMBOX.AddItem "55000001103         COSTA DE MARFIL - Persona Jurídica"
    COMBOX.AddItem "55000001138         EGIPTO - Persona Jurídica"
    COMBOX.AddItem "55000001146         ETIOPIA - Persona Jurídica"
    COMBOX.AddItem "55000001154         GABON - Persona Jurídica"
    COMBOX.AddItem "55000001162         GAMBIA - Persona Jurídica"
    COMBOX.AddItem "55000001170         GHANA - Persona Jurídica"
    COMBOX.AddItem "55000001189         GUINEA - Persona Jurídica"
    COMBOX.AddItem "55000001197         GUINEA ECUATORIAL - Persona Jurídica"
    COMBOX.AddItem "55000001200         KENIA - Persona Jurídica"
    COMBOX.AddItem "55000001219         LESOTHO - Persona Jurídica"
    COMBOX.AddItem "55000001227         REPUBLICA DE LIBERIA (Estado independiente) - Persona Jurídica"
    COMBOX.AddItem "55000001235         LIBIA - Persona Jurídica"
    COMBOX.AddItem "55000001243         MADAGASCAR - Persona Jurídica"
    COMBOX.AddItem "55000001278         MARRUECOS - Persona Jurídica"
    COMBOX.AddItem "55000001286         REPUBLICA DE MAURICIO - Persona Jurídica"
    COMBOX.AddItem "55000001294         MAURITANIA - Persona Jurídica"
    COMBOX.AddItem "55000001308         NIGER - Persona Jurídica"
    COMBOX.AddItem "55000001316         NIGERIA - Persona Jurídica"
    COMBOX.AddItem "55000001324         ZIMBABWE - Persona Jurídica"
    COMBOX.AddItem "55000001332         RUANDA - Persona Jurídica"
    COMBOX.AddItem "55000001340         SENEGAL - Persona Jurídica"
    COMBOX.AddItem "55000001359         SIERRA LEONA - Persona Jurídica"
    COMBOX.AddItem "55000001367         SOMALIA - Persona Jurídica"
    COMBOX.AddItem "55000001375         REINO DE SWAZILANDIA (Estado independiente) - Persona Jurídica"
    COMBOX.AddItem "55000001383         SUDAN - Persona Jurídica"
    COMBOX.AddItem "55000001405         TOGO - Persona Jurídica"
    COMBOX.AddItem "55000001413         REPUBLICA TUNECINA - Persona Jurídica"
    COMBOX.AddItem "55000001448         ZAMBIA - Persona Jurídica"
    COMBOX.AddItem "55000001456         POS.BRITANICA (AFRICA) - Persona Jurídica"
    COMBOX.AddItem "55000001464         POS.ESPAÑOLA (AFRICA) - Persona Jurídica"
    COMBOX.AddItem "55000001472         POS.FRANCESA (AFRICA) - Persona Jurídica"
    COMBOX.AddItem "55000001480         POS.PORTUGUESA (AFRICA) - Persona Jurídica"
    COMBOX.AddItem "55000001499         REPUBLICA DE ANGOLA - Persona Jurídica"
    COMBOX.AddItem "55000001502         REPUBLICA DE CABO VERDE (Estado independiente) - Persona Jurídica"
    COMBOX.AddItem "55000001510         MOZAMBIQUE - Persona Jurídica"
    COMBOX.AddItem "55000001529         CONGO REP.POPULAR - Persona Jurídica"
    COMBOX.AddItem "55000001537         CHAD - Persona Jurídica"
    COMBOX.AddItem "55000001545         MALAWI - Persona Jurídica"
    COMBOX.AddItem "55000001553         TANZANIA - Persona Jurídica"
    COMBOX.AddItem "55000001588         COSTA RICA - Persona Jurídica"
    COMBOX.AddItem "55000001618         ZAIRE - Persona Jurídica"
    COMBOX.AddItem "55000001626         BENIN - Persona Jurídica"
    COMBOX.AddItem "55000001634         MALI - Persona Jurídica"
    COMBOX.AddItem "55000001707         UGANDA - Persona Jurídica"
    COMBOX.AddItem "55000001715         SUDAFRICA, REP. DE - Persona Jurídica"
    COMBOX.AddItem "55000001812         REPUBLICA DE SEYCHELLES (Estado independiente) - Persona Jurídica"
    COMBOX.AddItem "55000001820         SANTO TOME Y PRINCIPE - Persona Jurídica"
    COMBOX.AddItem "55000001839         NAMIBIA - Persona Jurídica"
    COMBOX.AddItem "55000001847         GUINEA BISSAU - Persona Jurídica"
    COMBOX.AddItem "55000001855         ERITREA - Persona Jurídica"
    COMBOX.AddItem "55000001863         REPUBLICA DE DJIBUTI (Estado independiente) - Persona Jurídica"
    COMBOX.AddItem "55000001898         COMORAS - Persona Jurídica"
    COMBOX.AddItem "55000001987         INDETERMINADO (AFRICA) - Persona Jurídica"
    COMBOX.AddItem "55000002010         BARBADOS (Estado independiente) - Persona Jurídica"
    COMBOX.AddItem "55000002045         CANADA - Persona Jurídica"
    COMBOX.AddItem "55000002053         COLOMBIA - Persona Jurídica"
    COMBOX.AddItem "55000002096         DOMINICANA, REPUBLICA - Persona Jurídica"
    COMBOX.AddItem "55000002118         EL SALVADOR - Persona Jurídica"
    COMBOX.AddItem "55000002126         ESTADOS UNIDOS - Persona Jurídica"
    COMBOX.AddItem "55000002134         GUATEMALA - Persona Jurídica"
    COMBOX.AddItem "55000002142         REPUBLICA COOPERATIVA DE GUYANA (Estado independiente) - Persona Jurídica"
    COMBOX.AddItem "55000002150         HAITI - Persona Jurídica"
    COMBOX.AddItem "55000002169         HONDURAS - Persona Jurídica"
    COMBOX.AddItem "55000002177         JAMAICA - Persona Jurídica"
    COMBOX.AddItem "55000002185         MEXICO - Persona Jurídica"
    COMBOX.AddItem "55000002193         NICARAGUA - Persona Jurídica"
    COMBOX.AddItem "55000002207         REPUBLICA DE PANAMA (Estado independiente) - Persona Jurídica"
    COMBOX.AddItem "55000002215         ESTADO LIBRE ASOCIADO DE PUERTO RICO (Estado asoc. a EEUU) - Persona Jurídica"
    COMBOX.AddItem "55000002223         PERU - Persona Jurídica"
    COMBOX.AddItem "55000002258         ANTIGUA Y BARBUDA (Estado independiente) - Persona Jurídica"
    COMBOX.AddItem "55000002266         VENEZUELA - Persona Jurídica"
    COMBOX.AddItem "55000002274         POS.BRITANICA (AMERICA) - Persona Jurídica"
    COMBOX.AddItem "55000002282         POS.DANESA (AMERICA) - Persona Jurídica"
    COMBOX.AddItem "55000002290         POS.FRANCESA (AMERICA) - Persona Jurídica"
    COMBOX.AddItem "55000002304         POS.PAISES BAJOS (AMERICA) - Persona Jurídica"
    COMBOX.AddItem "55000002312         POS.E.E.U.U. (AMERICA) - Persona Jurídica"
    COMBOX.AddItem "55000002320         SURINAME - Persona Jurídica"
    COMBOX.AddItem "55000002339         EL COMMONWEALTH DE DOMINICA (Estado Asociado) - Persona Jurídica"
    COMBOX.AddItem "55000002347         SANTA LUCIA - Persona Jurídica"
    COMBOX.AddItem "55000002355         SAN VICENTE Y LAS GRANADINAS (Estado independiente) - Persona Jurídica"
    COMBOX.AddItem "55000002363         BELICE (Estado independiente) - Persona Jurídica"
    COMBOX.AddItem "55000002398         CUBA - Persona Jurídica"
    COMBOX.AddItem "55000002428         ECUADOR - Persona Jurídica"
    COMBOX.AddItem "55000002436         REPUBLICA DE TRINIDAD Y TOBAGO - Persona Jurídica"
    COMBOX.AddItem "55000002827         BUTAN - Persona Jurídica"
    COMBOX.AddItem "55000002843         MYANMAR (EX BIRMANIA) - Persona Jurídica"
    COMBOX.AddItem "55000002878         ISRAEL - Persona Jurídica"
    COMBOX.AddItem "55000002884         ESTADO ASOCIADO DE GRANADA (Estado independiente) - Persona Jurídica"
    COMBOX.AddItem "55000002894         FEDERACION DE SAN CRISTOBAL (Islas Saint Kitts and Nevis) - Persona Jurídica"
    COMBOX.AddItem "55000002908         COMUNIDAD DE LAS BAHAMAS (Estado independiente) - Persona Jurídica"
    COMBOX.AddItem "55000002916         TAILANDIA - Persona Jurídica"
    COMBOX.AddItem "55000002924         INDETERMINADO (AMERICA) - Persona Jurídica"
    COMBOX.AddItem "55000002932         IRAN - Persona Jurídica"
    COMBOX.AddItem "55000002983         ESTADO DE QATAR (Estado independiente) - Persona Jurídica"
    COMBOX.AddItem "55000003009         REINO HACHEMITA DE JORDANIA - Persona Jurídica"
    COMBOX.AddItem "55000003017         AFGANISTAN - Persona Jurídica"
    COMBOX.AddItem "55000003025         ARABIA SAUDITA - Persona Jurídica"
    COMBOX.AddItem "55000003033         ESTADO DE BAHREIN (Estado independiente) - Persona Jurídica"
    COMBOX.AddItem "55000003068         CAMBOYA (EX KAMPUCHEA) - Persona Jurídica"
    COMBOX.AddItem "55000003076         REPUBLICA DEMOCRATICA SOCIALISTA DE SRI LANKA - Persona Jurídica"
    COMBOX.AddItem "55000003084         COREA DEMOCRATICA  - Persona Jurídica"
    COMBOX.AddItem "55000003092         COREA REPUBLICANA - Persona Jurídica"
    COMBOX.AddItem "55000003106         CHINA REP.POPULAR - Persona Jurídica"
    COMBOX.AddItem "55000003114         REPUBLICA DE CHIPRE (Estado independiente) - Persona Jurídica"
    COMBOX.AddItem "55000003122         FILIPINAS - Persona Jurídica"
    COMBOX.AddItem "55000003130         TAIWAN - Persona Jurídica"
    COMBOX.AddItem "55000003149         GAZA - Persona Jurídica"
    COMBOX.AddItem "55000003157         INDIA - Persona Jurídica"
    COMBOX.AddItem "55000003165         INDONESIA - Persona Jurídica"
    COMBOX.AddItem "55000003173         IRAK - Persona Jurídica"
    COMBOX.AddItem "55000003203         JAPON - Persona Jurídica"
    COMBOX.AddItem "55000003238         ESTADO DE KUWAIT (Estado independiente) - Persona Jurídica"
    COMBOX.AddItem "55000003246         LAOS - Persona Jurídica"
    COMBOX.AddItem "55000003254         LIBANO - Persona Jurídica"
    COMBOX.AddItem "55000003262         MALASIA - Persona Jurídica"
    COMBOX.AddItem "55000003270         REPUBLICA DE MALDIVAS (Estado independiente) - Persona Jurídica"
    COMBOX.AddItem "55000003289         SULTANATO DE OMAN - Persona Jurídica"
    COMBOX.AddItem "55000003297         MONGOLIA - Persona Jurídica"
    COMBOX.AddItem "55000003300         NEPAL - Persona Jurídica"
    COMBOX.AddItem "55000003319         EMIRATOS ARABES UNIDOS (Estado independiente) - Persona Jurídica"
    COMBOX.AddItem "55000003327         PAKISTAN - Persona Jurídica"
    COMBOX.AddItem "55000003335         SINGAPUR - Persona Jurídica"
    COMBOX.AddItem "55000003343         SIRIA - Persona Jurídica"
    COMBOX.AddItem "55000003378         VIETNAM - Persona Jurídica"
    COMBOX.AddItem "55000003394         REPUBLICA DEL YEMEN - Persona Jurídica"
    COMBOX.AddItem "55000003416         POS.BRITANICA (HONG KONG) - Persona Jurídica"
    COMBOX.AddItem "55000003424         POS.JAPONESA (ASIA) - Persona Jurídica"
    COMBOX.AddItem "55000003440         MACAO - Persona Jurídica"
    COMBOX.AddItem "55000003459         BANGLADESH - Persona Jurídica"
    COMBOX.AddItem "55000003505         TURQUIA - Persona Jurídica"
    COMBOX.AddItem "55000003548         ITALIA - Persona Jurídica"
    COMBOX.AddItem "55000003556         TURKMENISTAN - Persona Jurídica"
    COMBOX.AddItem "55000003564         UZBEKISTAN - Persona Jurídica"
    COMBOX.AddItem "55000003572         TERRITORIOS AUTONOMOS PALESTINOS - Persona Jurídica"
    COMBOX.AddItem "55000003815         ISLANDIA - Persona Jurídica"
    COMBOX.AddItem "55000003882         GEORGIA - Persona Jurídica"
    COMBOX.AddItem "55000003890         TAYIKISTAN - Persona Jurídica"
    COMBOX.AddItem "55000003904         AZERBAIDZHAN - Persona Jurídica"
    COMBOX.AddItem "55000003912         BRUNEI DARUSSALAM (Estado independiente) - Persona Jurídica"
    COMBOX.AddItem "55000003920         KAZAJSTAN - Persona Jurídica"
    COMBOX.AddItem "55000003939         KIRGUISTAN - Persona Jurídica"
    COMBOX.AddItem "55000003963         INDETERMINADO (ASIA) - Persona Jurídica"
    COMBOX.AddItem "55000004013         REPUBLICA DE ALBANIA - Persona Jurídica"
    COMBOX.AddItem "55000004048         PRINCIPADO DEL VALLE DE ANDORRA - Persona Jurídica"
    COMBOX.AddItem "55000004056         AUSTRIA - Persona Jurídica"
    COMBOX.AddItem "55000004064         BELGICA - Persona Jurídica"
    COMBOX.AddItem "55000004072         BULGARIA - Persona Jurídica"
    COMBOX.AddItem "55000004099         DINAMARCA - Persona Jurídica"
    COMBOX.AddItem "55000004102         ESPAÑA - Persona Jurídica"
    COMBOX.AddItem "55000004110         FINLANDIA - Persona Jurídica"
    COMBOX.AddItem "55000004129         FRANCIA - Persona Jurídica"
    COMBOX.AddItem "55000004137         GRECIA - Persona Jurídica"
    COMBOX.AddItem "55000004145         HUNGRIA - Persona Jurídica"
    COMBOX.AddItem "55000004153         IRLANDA (EIRE) - Persona Jurídica"
    COMBOX.AddItem "55000004188         PRINCIPADO DE LIECHTENSTEIN (Estado independiente) - Persona Jurídica"
    COMBOX.AddItem "55000004196         GRAN DUCADO DE LUXEMBURGO - Persona Jurídica"
    COMBOX.AddItem "55000004218         PRINCIPADO DE MONACO - Persona Jurídica"
    COMBOX.AddItem "55000004226         NORUEGA - Persona Jurídica"
    COMBOX.AddItem "55000004234         PAISES BAJOS - Persona Jurídica"
    COMBOX.AddItem "55000004242         POLONIA - Persona Jurídica"
    COMBOX.AddItem "55000004250         PORTUGAL - Persona Jurídica"
    COMBOX.AddItem "55000004269         REINO UNIDO - Persona Jurídica"
    COMBOX.AddItem "55000004277         RUMANIA - Persona Jurídica"
    COMBOX.AddItem "55000004285         SERENISIMA REPUBLICA DE SAN MARINO (Estado independiente) - Persona Jurídica"
    COMBOX.AddItem "55000004293         SUECIA - Persona Jurídica"
    COMBOX.AddItem "55000004307         SUIZA - Persona Jurídica"
    COMBOX.AddItem "55000004315         SANTA SEDE (VATICANO) - Persona Jurídica"
    COMBOX.AddItem "55000004323         YUGOSLAVIA - Persona Jurídica"
    COMBOX.AddItem "55000004366         REPUBLICA DE MALTA (Estado independiente) - Persona Jurídica"
    COMBOX.AddItem "55000004382         ALEMANIA, REP. FED. - Persona Jurídica"
    COMBOX.AddItem "55000004390         BIELORUSIA - Persona Jurídica"
    COMBOX.AddItem "55000004404         ESTONIA - Persona Jurídica"
    COMBOX.AddItem "55000004412         LETONIA - Persona Jurídica"
    COMBOX.AddItem "55000004420         LITUANIA - Persona Jurídica"
    COMBOX.AddItem "55000004439         MOLDOVA - Persona Jurídica"
    COMBOX.AddItem "55000004463         BOSNIA HERZEGOVINA - Persona Jurídica"
    COMBOX.AddItem "55000004498         ESLOVENIA - Persona Jurídica"
    COMBOX.AddItem "55000004900         MACEDONIA - Persona Jurídica"
    COMBOX.AddItem "55000004919         POS.BRITANICA (EUROPA) - Persona Jurídica"
    COMBOX.AddItem "55000004986         INDETERMINADO (EUROPA) - Persona Jurídica"
    COMBOX.AddItem "55000004994         AUSTRALIA - Persona Jurídica"
    COMBOX.AddItem "55000005036         REPUBLICA DE NAURU (Estado independiente) - Persona Jurídica"
    COMBOX.AddItem "55000005044         NUEVA ZELANDA - Persona Jurídica"
    COMBOX.AddItem "55000005052         REPUBLICA DE VANUATU - Persona Jurídica"
    COMBOX.AddItem "55000005069         SAMOA OCCIDENTAL - Persona Jurídica"
    COMBOX.AddItem "55000005079         POS.AUSTRALIANA (OCEANIA) - Persona Jurídica"
    COMBOX.AddItem "55000005087         POS.BRITANICA (OCEANIA) - Persona Jurídica"
    COMBOX.AddItem "55000005095         POS.FRANCESA (OCEANIA) - Persona Jurídica"
    COMBOX.AddItem "55000005109         POS.NEOCELANDESA (OCEANIA) - Persona Jurídica"
    COMBOX.AddItem "55000005117         POS.E.E.U.U. (OCEANIA) - Persona Jurídica"
    COMBOX.AddItem "55000005125         FIJI, ISLAS - Persona Jurídica"
    COMBOX.AddItem "55000005133         PAPUA, ISLAS - Persona Jurídica"
    COMBOX.AddItem "55000005168         KIRIBATI - Persona Jurídica"
    COMBOX.AddItem "55000005176         TUVALU - Persona Jurídica"
    COMBOX.AddItem "55000005184         ISLAS SALOMON - Persona Jurídica"
    COMBOX.AddItem "55000005192         REINO DE TONGA (Estado independiente) - Persona Jurídica"
    COMBOX.AddItem "55000005206         REPUBLICA DE LAS ISLAS MARSHALL (Estado independiente) - Persona Jurídica"
    COMBOX.AddItem "55000005214         ISLAS MARIANAS - Persona Jurídica"
    COMBOX.AddItem "55000005907         MICRONESIA ESTADOS FED. - Persona Jurídica"
    COMBOX.AddItem "55000005915         PALAU - Persona Jurídica"
    COMBOX.AddItem "55000005982         INDETERMINADO (OCEANIA) - Persona Jurídica"
    COMBOX.AddItem "55000006016         RUSA, FEDERACION - Persona Jurídica"
    COMBOX.AddItem "55000006024         ARMENIA - Persona Jurídica"
    COMBOX.AddItem "55000006032         CROACIA - Persona Jurídica"
    COMBOX.AddItem "55000006040         UCRANIA - Persona Jurídica"
    COMBOX.AddItem "55000006067         ESLOVACA, REPUBLICA - Persona Jurídica"
    COMBOX.AddItem "55000006520         ANGUILA (Territorio no autónomo del Reino Unido) - Persona Jurídica"
    COMBOX.AddItem "55000006539         ARUBA (Territorio de Países Bajos) - Persona Jurídica"
    COMBOX.AddItem "55000006547         ISLAS DE COOK (Territorio autónomo asociado a Nueva Zelanda) - Persona Jurídica"
    COMBOX.AddItem "55000006555         PATAU - Persona Jurídica"
    COMBOX.AddItem "55000006563         POLINESIA FRANCESA (Territorio de Ultramar de Francia) - Persona Jurídica"
    COMBOX.AddItem "55000006598         ANTILLAS HOLANDESAS (Territorio de Países Bajos) - Persona Jurídica"
    COMBOX.AddItem "55000006628         ASCENCION - Persona Jurídica"
    COMBOX.AddItem "55000006636         BERMUDAS (Territorio no autónomo del Reino Unido) - Persona Jurídica"
    COMBOX.AddItem "55000006644         CAMPIONE D@ITALIA - Persona Jurídica"
    COMBOX.AddItem "55000006652         COLONIA DE GIBRALTAR - Persona Jurídica"
    COMBOX.AddItem "55000006660         GROENLANDIA - Persona Jurídica"
    COMBOX.AddItem "55000006679         GUAM (Territorio no autónomo de los EEUU) - Persona Jurídica"
    COMBOX.AddItem "55000006687         HONK KONG (Territorio de China) - Persona Jurídica"
    COMBOX.AddItem "55000006695         ISLAS AZORES - Persona Jurídica"
    COMBOX.AddItem "55000006709         ISLAS DEL CANAL:Guernesey,Jersey,Alderney,G.Stark,L.Sark,etc - Persona Jurídica"
    COMBOX.AddItem "55000006717         ISLAS CAIMAN (Territorio no autónomo del Reino Unido) - Persona Jurídica"
    COMBOX.AddItem "55000006725         ISLA CHRISTMAS - Persona Jurídica"
    COMBOX.AddItem "55000006733         ISLA DE COCOS O KEELING - Persona Jurídica"
    COMBOX.AddItem "55000006768         ISLA DE MAN (Territorio del Reino Unido) - Persona Jurídica"
    COMBOX.AddItem "55000006776         ISLA DE NORFOLK - Persona Jurídica"
    COMBOX.AddItem "55000006784         ISLAS TURKAS Y CAICOS (Territorio no autónomo del R. Unido) - Persona Jurídica"
    COMBOX.AddItem "55000006792         ISLAS PACIFICO - Persona Jurídica"
    COMBOX.AddItem "55000006806         ISLA DE SAN PEDRO Y MIGUELON - Persona Jurídica"
    COMBOX.AddItem "55000006814         ISLA QESHM - Persona Jurídica"
    COMBOX.AddItem "55000006822         ISLAS VIRGENES BRITANICAS(Territorio no autónomo de R.UNIDO) - Persona Jurídica"
    COMBOX.AddItem "55000006830         ISLAS VIRGENES DE ESTADOS UNIDOS DE AMERICA - Persona Jurídica"
    COMBOX.AddItem "55000006849         LABUAN - Persona Jurídica"
    COMBOX.AddItem "55000006857         MADEIRA (Territorio de Portugal) - Persona Jurídica"
    COMBOX.AddItem "55000006865         MONTSERRAT (Territorio no autónomo del Reino Unido) - Persona Jurídica"
    COMBOX.AddItem "55000006873         NIUE - Persona Jurídica"
    COMBOX.AddItem "55000006903         PITCAIRN - Persona Jurídica"
    COMBOX.AddItem "55000006938         REGIMEN APLICABLE A LAS SA FINANCIERAS(ley 11.073 de la ROU) - Persona Jurídica"
    COMBOX.AddItem "55000006946         SANTA ELENA - Persona Jurídica"
    COMBOX.AddItem "55000006954         SAMOA AMERICANA (Territorio no autónomo de los EEUU) - Persona Jurídica"
    COMBOX.AddItem "55000006962         ARCHIPIELAGO DE SVBALBARD - Persona Jurídica"
    COMBOX.AddItem "55000006970         TRISTAN DA CUNHA - Persona Jurídica"
    COMBOX.AddItem "55000006989         TRIESTE (Italia) - Persona Jurídica"
    COMBOX.AddItem "55000006997         TOKELAU - Persona Jurídica"
    COMBOX.AddItem "55000007004         ZONA LIBRE DE OSTRAVA (ciudad de la antigua Checoeslovaquia) - Persona Jurídica"
    COMBOX.AddItem "55000009988         PARA PERSONAS FISICAS DE INDETERMINADO (CONTINENTE) - Persona Jurídica"
    COMBOX.AddItem "55000009996         PARA PERSONAS FISICAS DE OTROS PAISES - Persona Jurídica"


End Sub
Sub CARGACOMBO5(COMBOX As ComboBox)

    COMBOX.AddItem ""
    COMBOX.AddItem "EXW"
    COMBOX.AddItem "FCA"
    COMBOX.AddItem "FAS"
    COMBOX.AddItem "FOB"
    COMBOX.AddItem "CFR"
    COMBOX.AddItem "CIF"
    COMBOX.AddItem "CPT"
    COMBOX.AddItem "CIP"
    COMBOX.AddItem "DAF"
    COMBOX.AddItem "DES"
    COMBOX.AddItem "DEQ"
    COMBOX.AddItem "DDU"
    COMBOX.AddItem "DDP"

End Sub
Sub CARGACOMBO6(COMBOX As ComboBox)

    COMBOX.AddItem ""
    COMBOX.AddItem "1 Español"
    COMBOX.AddItem "2 Inglés"
    COMBOX.AddItem "3 Portugués"
    
End Sub

Private Sub grabar_Click()
   Call Command2_Click
End Sub

Private Sub Label25_DblClick()
   
    Call SETULTREG&("!COMPASOC", 0)
    JJ& = 0
    For JX% = 0 To List1.ListCount - 1
      L1$ = List1.List(JX%)
      TIPCO% = Left$(TRIM$(L1$), 2)
      PVTA$ = TRIM$(Mid$(L1$, 20, 4))
      NCBTE$ = TRIM$(Mid$(L1$, 32, 8))
      '
      YY$ = REGBLAN$("!COMPASOC")
      Call REPLA(YY$, MKI$(TIPCO%), 1, 2)
      Call REPLA(YY$, PVTA$, 3, 4)
      Call REPLA(YY$, NCBTE$, 7, 8)
      JJ& = JJ& + 1
      Call GRAREG("!COMPASOC", YY$, JJ&)
      '
    Next JX%

   
10 VTB% = VENTABU%("COMPASOC")
   If VTB% < 0 Then Exit Sub
   '
   'VALIDO
   JJ& = 0
   List1.Clear
   
   For U& = 1 To ULTREG&("!COMPASOC")
         XX$ = REGLEIDO$("!COMPASOC", U&)
         TIPCO% = CVI(Mid$(XX$, 1, 2))
         DESCRI$ = TRIM$(ECOARCH$("TATICOMP", Chr$(TIPCO%)))
         PVTA$ = TRIM$(Mid$(XX$, 3, 4))
         NCBTE$ = TRIM$(Mid$(XX$, 7, 8))
         If TIPCO% > 0 Then
            If XVALO(PVTA$) < 1 Then Call COMUNI("Punto de Venta No Válido"): GoTo 10
            If XVALO(NCBTE$) < 1 Then Call COMUNI("Número de Comprobante No Válido"): GoTo 10
         End If
    Next U&
    '
    List1.Clear
    For U& = 1 To ULTREG&("!COMPASOC")
         XX$ = REGLEIDO$("!COMPASOC", U&)
         TIPCO% = CVI(Mid$(XX$, 1, 2))
         DESCRI$ = TRIM$(ECOARCH$("TATICOMP", Chr$(TIPCO%)))
         PVTA$ = TRIM$(Mid$(XX$, 3, 4))
         While Len(PVTA$) < 4: PVTA$ = "0" & PVTA$: Wend
         NCBTE$ = TRIM$(Mid$(XX$, 7, 8))
         While Len(NCBTE$) < 8: NCBTE$ = "0" & NCBTE$: Wend
         Call REPLA(TX$, TRIM$(Str$(TIPCO%)) & " " & DESCRI$, 1, 18)
         Call REPLA(TX$, PVTA$, 20, 4)
         Call REPLA(TX$, NCBTE$, 32, 8)
         List1.AddItem TX$
     Next U&
     '
     Call CIERRARCH("!COMPASOC")
     Call CIERRARCH("COMPASOC")

      
End Sub

Private Sub List1_Click()
   If List1.ListCount > 0 Then
      Call Label25_DblClick
   End If
End Sub


Private Sub salir_Click()
   Call Command5_Click
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call Command4_Click
   End If
End Sub


