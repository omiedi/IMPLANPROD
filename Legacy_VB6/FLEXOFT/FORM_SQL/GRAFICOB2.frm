VERSION 5.00
Object = "{02B5E320-7292-11CF-93D5-0020AF99504A}#1.0#0"; "MSCHART.OCX"
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
   Begin MSChartLib.MSChart MSChart1 
      Height          =   4005
      Left            =   120
      OleObjectBlob   =   "GRAFICOB.frx":0000
      TabIndex        =   0
      Top             =   120
      Width           =   11700
   End
End
Attribute VB_Name = "GRAFICOB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
