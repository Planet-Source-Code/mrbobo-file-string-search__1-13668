VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find files containing string..."
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   5730
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox rtftext 
      Height          =   255
      Left            =   960
      TabIndex        =   8
      Top             =   3000
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form1.frx":0000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   2640
      Width           =   2295
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   3120
      TabIndex        =   4
      Top             =   840
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   360
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2640
      Width           =   2535
   End
   Begin VB.DirListBox Dir1 
      Height          =   1665
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   2535
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3120
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   3120
      TabIndex        =   9
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Text to find"
      Height          =   255
      Left            =   3120
      TabIndex        =   7
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Filter"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   2400
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim filter As String
Private Sub Combo1_Click()
'set the filter variable
Select Case Combo1.ListIndex
Case 0
    filter = "txt"
Case 1
    filter = "doc"
Case 2
    filter = ""
End Select
End Sub

Private Sub Command1_Click()
'Ok lets go
If Text1.Text = "" Then
    MsgBox "Nothing to search for !!", vbExclamation, "Bobo Enterprises"
    Exit Sub
End If
List1.Clear 'Remove old search results
GetFiles Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
'fill the combo with file types for filter
Combo1.AddItem "Text Files (*.txt)"
Combo1.AddItem "Document Files (*.doc)"
Combo1.AddItem "All Files (*.*)"
Combo1.ListIndex = 0
End Sub
Private Sub GetFiles(MyDir As String)
'Standard Dir function applying filters and string search if required
Dim temp As String
Dim textfound As Integer
Dim myext As String
If Right(MyDir, 1) <> "\" Then MyDir = MyDir + "\"
temp = Dir(MyDir)
Do While temp <> ""
    If temp <> "." And temp <> ".." Then
        If GetAttr(MyDir + temp) <> vbDirectory Then 'If its a file
            If filter = "" Then 'filter is All files
                rtftext.LoadFile MyDir + temp 'load file
                textfound = rtftext.Find(Text1.Text) 'look for string
                If textfound <> -1 Then List1.AddItem temp 'if found add to list
            Else 'filter is txt or doc files
                myext = Mid$(temp, InStrRev(temp, ".") + 1) 'extension only
                If myext = filter Then 'If file is txt or doc
                    rtftext.LoadFile MyDir + temp
                    textfound = rtftext.Find(Text1.Text)
                    If textfound <> -1 Then List1.AddItem temp
                End If
            End If
        End If
    End If
    temp = Dir
    Label3 = "Files found " + Str(List1.ListCount)
    Label3.Refresh
Loop
End Sub
