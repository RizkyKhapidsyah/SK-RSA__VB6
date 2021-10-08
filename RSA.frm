VERSION 5.00
Begin VB.Form RSA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RSA Demonstration"
   ClientHeight    =   6450
   ClientLeft      =   4485
   ClientTop       =   2100
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDecrypt 
      Caption         =   "String Decrypt"
      Height          =   345
      Left            =   2265
      TabIndex        =   20
      Top             =   6030
      Width           =   1335
   End
   Begin VB.TextBox Text10 
      Height          =   1095
      Left            =   45
      MultiLine       =   -1  'True
      TabIndex        =   19
      Top             =   4815
      Width           =   4470
   End
   Begin VB.TextBox Text9 
      Height          =   1095
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   18
      Top             =   3480
      Width           =   4470
   End
   Begin VB.TextBox Text3 
      Height          =   555
      Left            =   60
      TabIndex        =   17
      Text            =   "Hello World"
      Top             =   2685
      Width           =   4470
   End
   Begin VB.CommandButton cmdEncrypt 
      Caption         =   "String Encrypt"
      Height          =   345
      Left            =   555
      TabIndex        =   16
      Top             =   6030
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Create - (Public - Private Keys)"
      Height          =   2370
      Left            =   60
      TabIndex        =   0
      Top             =   45
      Width           =   4470
      Begin VB.TextBox txtPhi 
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   11
         ToolTipText     =   "Phi = (P-1) * ( Q-1)"
         Top             =   1095
         Width           =   1005
      End
      Begin VB.TextBox txtQ 
         Height          =   315
         Left            =   1485
         Locked          =   -1  'True
         TabIndex        =   10
         ToolTipText     =   "Random number Q"
         Top             =   495
         Width           =   1005
      End
      Begin VB.TextBox txtP 
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   9
         ToolTipText     =   "Random number P"
         Top             =   495
         Width           =   1005
      End
      Begin VB.TextBox txtD 
         Height          =   315
         Left            =   2850
         Locked          =   -1  'True
         TabIndex        =   7
         ToolTipText     =   "D = Inverse of E"
         Top             =   1920
         Width           =   1005
      End
      Begin VB.TextBox txtN 
         Height          =   315
         Left            =   1485
         Locked          =   -1  'True
         TabIndex        =   4
         ToolTipText     =   "N = P * Q"
         Top             =   1920
         Width           =   1005
      End
      Begin VB.TextBox txtE 
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   3
         ToolTipText     =   "E = Random number relatively prime to PHI"
         Top             =   1920
         Width           =   1005
      End
      Begin VB.CommandButton cmdKeyGen 
         Caption         =   "&Generate Keys"
         Height          =   360
         Left            =   3000
         TabIndex        =   2
         Top             =   225
         Width           =   1335
      End
      Begin VB.Label Label15 
         Caption         =   "N  =  P * Q"
         Height          =   255
         Left            =   1485
         TabIndex        =   15
         Top             =   1200
         Width           =   1635
      End
      Begin VB.Label Label14 
         Caption         =   "Phi  =  (P-1) * (Q-1)"
         Height          =   255
         Left            =   1485
         TabIndex        =   14
         Top             =   945
         Width           =   1635
      End
      Begin VB.Label Label13 
         Caption         =   "(Phi)"
         Height          =   270
         Left            =   120
         TabIndex        =   13
         Top             =   825
         Width           =   885
      End
      Begin VB.Label Label12 
         Caption         =   "(Q)"
         Height          =   270
         Left            =   1485
         TabIndex        =   12
         Top             =   225
         Width           =   885
      End
      Begin VB.Label Label10 
         Caption         =   "'D' Decrypter key"
         Height          =   195
         Left            =   2835
         TabIndex        =   8
         Top             =   1680
         Width           =   1485
      End
      Begin VB.Label Label3 
         Caption         =   "'N' Modulus"
         Height          =   225
         Left            =   1485
         TabIndex        =   6
         Top             =   1665
         Width           =   1590
      End
      Begin VB.Label Label2 
         Caption         =   "'E' Encrypter key"
         Height          =   225
         Left            =   120
         TabIndex        =   5
         Top             =   1665
         Width           =   1260
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         X1              =   30
         X2              =   4455
         Y1              =   1530
         Y2              =   1530
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         X1              =   15
         X2              =   4440
         Y1              =   1515
         Y2              =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "(P)"
         Height          =   270
         Left            =   135
         TabIndex        =   1
         Top             =   225
         Width           =   885
      End
   End
   Begin VB.Label Label6 
      Caption         =   "Decrypted text:"
      Height          =   240
      Left            =   60
      TabIndex        =   23
      Top             =   4590
      Width           =   1920
   End
   Begin VB.Label Label5 
      Caption         =   "Encrypted text:"
      Height          =   225
      Left            =   60
      TabIndex        =   22
      Top             =   3255
      Width           =   1740
   End
   Begin VB.Label Label4 
      Caption         =   "Text to encrypt:"
      Height          =   210
      Left            =   75
      TabIndex        =   21
      Top             =   2460
      Width           =   1815
   End
   Begin VB.Menu mnuExit 
      Caption         =   "&Exit"
   End
   Begin VB.Menu mnuView 
      Caption         =   "&Information"
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "RSA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim E, D, N As Double

Private Sub cmdDecrypt_Click()
Text10 = dec(Text9, key(2), key(3))
End Sub

Private Sub cmdEncrypt_Click()
Text9 = enc(Text3, key(1), key(3))
End Sub

Private Sub cmdKeyGen_Click()
keyGen
                        txtP = p      'P
                        txtQ = q      'Q
                        txtPhi = PHI  'PHI
                        txtE = key(1) 'E
                        txtD = key(2) 'D
                        txtN = key(3) 'N
                        Text4 = txtE
                        Text5 = txtN
                        Text6 = txtN
                        Text7 = txtD
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show vbModal
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuView_Click()
frmView.Show vbModal
End Sub
