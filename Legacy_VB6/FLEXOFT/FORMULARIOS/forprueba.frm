VERSION 5.00
Begin VB.Form forprueba 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   495
      Left            =   3720
      TabIndex        =   3
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   615
      Left            =   3240
      TabIndex        =   2
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   960
      TabIndex        =   0
      Top             =   720
      Width           =   2295
   End
End
Attribute VB_Name = "forprueba"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  frmforcarcli.Show
End Sub

Private Sub Command2_Click()
Shell "c:\Archivos de programa\Devstudio\VB\Report\CRW32.exe c:\clientes\sabo_win\inforpru.rpt", vbMaximizedFocus
End Sub

Private Sub Command3_Click()
    Dim dbclientes As Database

Set dbclientes = OpenDatabase("C:\CLIENTES\SABO_WIN\PRUEBA EN ACCESS.mdb")
Set tdfNuevo = dbclientes.TableDefs("Articulos")

    With tdfNuevo
        ' Crea los campos y los agrega al objeto
        ' TableDef nuevo. Debe hacer esto después de
        ' anexar el objeto TableDef a la colección
        ' TableDefs de la base de datos Neptuno.
        On Error Resume Next
        .Fields.Append .CreateField("Nombre", dbText)

        .Fields.Append .CreateField("Teléfono", dbText)
        .Fields.Append .CreateField("Notas", dbMemo)
        On Error GoTo 0
    End With


    dbclientes.Close

End Sub

Private Sub Command4_Click()
frmMaster.Show

End Sub
