VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00404040&
   Caption         =   "Form2"
   ClientHeight    =   8484
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   11880
   LinkTopic       =   "Form2"
   ScaleHeight     =   8484
   ScaleWidth      =   11880
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Enter  "
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      TabIndex        =   1
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Caption         =   "B.MANIKANDAN, K.NAVANEETHAN, I.PRAVEEN KUMAR"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   26.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2412
      Index           =   2
      Left            =   3240
      TabIndex        =   3
      Top             =   5640
      Width           =   5412
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Caption         =   "Play Hide And Seek With High Tech....."
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   26.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Index           =   1
      Left            =   1560
      TabIndex        =   2
      Top             =   1800
      Width           =   11652
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Caption         =   "STEGANOGRAPHY IN NETWORK SECURITY"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   26.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Top             =   4920
      Width           =   11652
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Form2
Form2.Visible = False
Load Form1
Form1.Visible = True
End Sub

