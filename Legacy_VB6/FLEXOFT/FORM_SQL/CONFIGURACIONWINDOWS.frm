VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5490
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13140
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   13140
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command11 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2640
      TabIndex        =   10
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2640
      TabIndex        =   9
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2640
      TabIndex        =   8
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1320
      TabIndex        =   7
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1320
      TabIndex        =   6
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1320
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command1"
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command1"
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command1"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim O_Registro As WshShell
  
  
Private Const LocalMachine As String = "HKEY_LOCAL_MACHINE"
Private Const CurrentUser As String = "HKEY_CURRENT_USER"
  
'------------------------------------------
  
'    Configuración de la MONEDA
  
'------------------------------------------
  
' Número de dígitos decimales para la moneda
Private Sub Command1_Click()
MsgBox Leer_Dato(CurrentUser, "iCurrDigits"), vbInformation
End Sub
  
  
  
'Símbolo de Moneda
Private Sub Command2_Click()
    MsgBox Leer_Dato(CurrentUser, "sCurrency"), vbInformation
End Sub
  
'Símbolo del separador decimal para la moneda
Private Sub Command3_Click()
    MsgBox Leer_Dato(CurrentUser, "sMonDecimalSep"), vbInformation
End Sub
  
'Símbolo del separador de miles para la moneda
Private Sub Command4_Click()
    MsgBox Leer_Dato(CurrentUser, "sMonThousandSep"), vbInformation
End Sub
  
  
  
'------------------------------------------
  
'    Configuración de los NUMEROS
  
'------------------------------------------
  
' Número de dígitos decimales
Private Sub Command5_Click()
    MsgBox Leer_Dato(CurrentUser, "iDigits"), vbInformation
End Sub
  
'Símbolo del separador de miles
Private Sub Command6_Click()
    MsgBox Leer_Dato(CurrentUser, "sThousand"), vbInformation
End Sub
  
'Símbolo del separador decimal
Private Sub Command7_Click()
    MsgBox Leer_Dato(CurrentUser, "sDecimal"), vbInformation
End Sub
  
'------------------------------------------
  
'    Configuración de HORA y FECHA
  
'------------------------------------------
  
' Formato de hora
Private Sub Command8_Click()
    MsgBox Leer_Dato(CurrentUser, "sTimeFormat"), vbInformation
End Sub
  
'Símbolo separador de HORA
Private Sub Command9_Click()
    MsgBox Leer_Dato(CurrentUser, "sTime"), vbInformation
End Sub
  
  
' Formato de Fecha Corta
Private Sub Command10_Click()
    MsgBox Leer_Dato(CurrentUser, "sShortDate"), vbInformation
End Sub
  
' Formato de Fecha Larga
Private Sub Command11_Click()
    MsgBox Leer_Dato(CurrentUser, "sLongDate"), vbInformation
End Sub
  
  
  
' Función que lee el valor del registro en la rama Control Panel\International
  
Function Leer_Dato(Principal As String, Valor As String) As String
    Leer_Dato = O_Registro.RegRead(Principal & "\Control Panel\International\" & Valor)
End Function
  
  
  
Private Sub Form_Load()
    Set O_Registro = New WshShell
  
    Command1.Caption = " N° Dígitos decimales"
    Command2.Caption = " Símbolo de Moneda"
    Command3.Caption = " Separador decimal"
    Command4.Caption = " Separador de miles"
    Command5.Caption = " N° Dígitos decimales"
    Command6.Caption = " Separador de miles"
    Command7.Caption = " Separador decimal"
    Command8.Caption = " Formato de hora "
    Command9.Caption = " Separador de HORA"
    Command10.Caption = " Fecha Corta "
    Command11.Caption = " Fecha Larga"
  
End Sub
  
Private Sub Form_Unload(Cancel As Integer)
    Set O_Registro = Nothing
End Sub
