VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H80000008&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Id3"
   ClientHeight    =   2715
   ClientLeft      =   8940
   ClientTop       =   5430
   ClientWidth     =   5760
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      Caption         =   "Song:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   2175
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5175
      Begin VB.TextBox Text4 
         BackColor       =   &H8000000A&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H8000000A&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   1080
         Width           =   3615
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H8000000A&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   720
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000A&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Date:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Album:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Title:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Artist:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
ID3.GetTag Form1.List1.Text
If ID3.HTag = True Then
    Text1.Locked = False
    Text2.Locked = False
    Text3.Locked = False
    Text4.Locked = False
    Text1.Text = ID3.Artist
    Text2.Text = ID3.Title
    Text3.Text = ID3.Album
    Text4.Text = ID3.Dated
    Text1.Locked = True
    Text2.Locked = True
    Text3.Locked = True
    Text4.Locked = True
Else
    Text1.Locked = False
    Text2.Locked = False
    Text3.Locked = False
    Text4.Locked = False
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text1.Locked = True
    Text2.Locked = True
    Text3.Locked = True
    Text4.Locked = True
End If
End Sub
