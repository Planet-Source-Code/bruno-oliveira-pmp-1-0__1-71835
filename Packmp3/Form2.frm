VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "                                                  .:Logon:."
   ClientHeight    =   3375
   ClientLeft      =   3615
   ClientTop       =   1080
   ClientWidth     =   5130
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Sair"
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Entrar"
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2040
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Utilizador"
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "administrador" And Text2.Text = "rtypz" Then
Load Form1
Form1.Show
Unload Form2
Form2.Hide
Else
If Text1.Text = "Brunão" And Text2.Text = "xxl90" Then
Load Form1
Form1.Show
Unload Form2
Form2.Hide
Else
resp = MsgBox("O utilizador não tem acesso a aplicação", 32, "Erro:")
End If
End If
End Sub

Private Sub Command2_Click()
End
End Sub

