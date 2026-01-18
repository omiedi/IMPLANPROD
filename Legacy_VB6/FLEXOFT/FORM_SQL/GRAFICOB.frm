VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.ocx"
Begin VB.Form GRAFICOB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estadística de Cobranza por Cliente"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   11925
   StartUpPosition =   2  'CenterScreen
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   3975
      Left            =   120
      OleObjectBlob   =   "GRAFICOB.frx":0000
      TabIndex        =   0
      Top             =   120
      Width           =   11775
   End
End
Attribute VB_Name = "GRAFICOB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
