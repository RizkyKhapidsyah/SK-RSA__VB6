VERSION 5.00
Begin VB.Form frmView 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RSA Information and Source Code"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   360
      Left            =   2265
      TabIndex        =   1
      Top             =   5340
      Width           =   1170
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000000&
      ForeColor       =   &H00000000&
      Height          =   5175
      Left            =   75
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmView.frx":0000
      Top             =   75
      Width           =   6030
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
Unload Me
End Sub
