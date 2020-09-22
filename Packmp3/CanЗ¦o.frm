VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Song:"
   ClientHeight    =   3690
   ClientLeft      =   9255
   ClientTop       =   5115
   ClientWidth     =   5490
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      Caption         =   "Description:"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   1800
         TabIndex        =   8
         Top             =   2520
         Width           =   2295
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Size:"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000E&
         Height          =   615
         Left            =   1680
         TabIndex        =   6
         Top             =   1920
         Width           =   3375
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1440
         TabIndex        =   5
         Top             =   1560
         Width           =   3255
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1440
         TabIndex        =   4
         Top             =   1080
         Width           =   3135
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Localization:"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Duration:"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   1080
         Width           =   1695
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
On Error Resume Next
tsec = Form7.MediaPlayer1.Duration
Dim st, mt As Integer
mt = tsec \ 60
st = Round(tsec - (mt * 60))
If st < 10 Then
Label5.Caption = mt & ":0" & st
Else
If st > 10 Then
Label5.Caption = mt & ":" & st
End If
If sec > 10 Then
Label5.Caption = mt & ":0" & st
End If
End If
If st >= 10 Then
Label5.Caption = mt & ":" & st
End If
Label4.Caption = Form1.Label10.Caption
Label6.Caption = Form7.MediaPlayer1.FileName
If Not Form7.MediaPlayer1.FileName = "" Then
Label8.Caption = FileLen(Form7.MediaPlayer1.FileName) \ 1000 & "Kb's"
End If
End Sub




