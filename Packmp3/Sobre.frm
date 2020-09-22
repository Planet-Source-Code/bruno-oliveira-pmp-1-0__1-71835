VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   3135
   ClientLeft      =   8775
   ClientTop       =   5235
   ClientWidth     =   6495
   Icon            =   "Sobre.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   6495
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000A&
         Caption         =   "Packmp3"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1695
         Left            =   2760
         TabIndex        =   1
         Top             =   360
         Width           =   3255
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Copyright © 2008"
            Height          =   255
            Left            =   240
            TabIndex        =   5
            Top             =   360
            Width           =   2775
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "by  Bruno Capela Oliveira"
            Height          =   255
            Left            =   240
            TabIndex        =   4
            Top             =   600
            Width           =   2775
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "E-mails:"
            Height          =   255
            Left            =   240
            TabIndex        =   3
            Top             =   840
            Width           =   2775
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "b.capela.oliveira@gmail.com"
            Height          =   375
            Left            =   240
            TabIndex        =   2
            Top             =   1080
            Width           =   2775
         End
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   2055
         Left            =   240
         Picture         =   "Sobre.frx":044A
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2175
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib _
        "shell32.dll" Alias "ShellExecuteA" _
        (ByVal hwnd As Long, ByVal lpOperation _
        As String, ByVal lpFile As String, _
        ByVal lpParameters As String, ByVal _
        lpDirectory As String, ByVal nShowCmd _
        As Long) As Long

Private Sub Label4_Click()
Call ShellExecute(0&, vbNullString, "mailto:" & _
       Label4.Caption, vbNullString, vbNullString, _
       SW_SHOWNORMAL)
End Sub

Private Sub Label5_Click()
Call ShellExecute(0&, vbNullString, "mailto:" & _
       Label5.Caption, vbNullString, vbNullString, _
       SW_SHOWNORMAL)
End Sub
