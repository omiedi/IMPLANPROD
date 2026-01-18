VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Begin VB.Form PROBOLRE 
   BackColor       =   &H00FFFF80&
   Caption         =   "Procesamiento Boletas de Recepción"
   ClientHeight    =   1305
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   5850
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1305
   ScaleWidth      =   5850
   StartUpPosition =   1  'CenterOwner
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   165
      Left            =   1155
      TabIndex        =   0
      Top             =   840
      Width           =   4530
      _ExtentX        =   7990
      _ExtentY        =   291
      _Version        =   327680
      Appearance      =   1
      MouseIcon       =   "PROBOLRE.frx":0000
   End
   Begin VB.Image Image1 
      Height          =   645
      Left            =   210
      Picture         =   "PROBOLRE.frx":001C
      Stretch         =   -1  'True
      Top             =   315
      Width           =   540
   End
End
Attribute VB_Name = "PROBOLRE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
