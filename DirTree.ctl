VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl DirTree 
   BackStyle       =   0  'Transparent
   ClientHeight    =   4446
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3354
   KeyPreview      =   -1  'True
   ScaleHeight     =   342
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   258
   Begin VB.DirListBox SubDirs 
      Height          =   273
      Left            =   494
      TabIndex        =   3
      Top             =   2522
      Visible         =   0   'False
      Width           =   741
   End
   Begin VB.DriveListBox Drives 
      Height          =   273
      Left            =   1287
      TabIndex        =   2
      Top             =   1534
      Visible         =   0   'False
      Width           =   637
   End
   Begin VB.DirListBox Dirs 
      Height          =   468
      Left            =   1248
      TabIndex        =   1
      Top             =   1937
      Visible         =   0   'False
      Width           =   663
   End
   Begin MSComctlLib.TreeView GenView 
      Height          =   4212
      Left            =   78
      TabIndex        =   0
      Top             =   104
      Width           =   3159
      _ExtentX        =   5583
      _ExtentY        =   7428
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      Style           =   7
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList icons 
      Left            =   1001
      Top             =   936
      _ExtentX        =   911
      _ExtentY        =   911
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DirTree.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DirTree.ctx":0112
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "DirTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim TempDir As String

'control initialization
Private Sub UserControl_Initialize()

'set initial dir
Dirs.path = "c:\"
SubDirs.path = "c:\"
Dim i As Integer
GenView.ImageList = icons

icons.ListImages.Add , , GetIcon("a:\")
GenView.Nodes.Add , , "MyComputer", "My Computer", 1

'floppy
'GenView.Nodes.Add "MyComputer", tvwChild, "a:", "a:", 3
For i = 1 To Drives.ListCount - 1
    icons.ListImages.Add , , GetIcon(Mid(Drives.List(i), 1, 2) & "\") 'get associated icon
    GenView.Nodes.Add "MyComputer", tvwChild, Mid(Drives.List(i), 1, 2), Drives.List(i), i + 3
    If CountSubDirs(Drives.List(i)) > 0 Then
        GenView.Nodes.Add Mid(Drives.List(i), 1, 2), tvwChild, , "dummy"
    End If
Next
GenView.Nodes.Item(1).Expanded = True
End Sub


Private Sub GenView_Expand(ByVal Node As MSComctlLib.Node)

Dim relation As String
Dim ParentPath As String

If Node.Child.text = "dummy" Then
    GenView.Nodes.Remove (Node.Child.Index)
End If
If Not DirExist(Node.key) Then Exit Sub
relation = Node.key
Dirs.path = Node.key
  Dim i As Integer
    For i = 0 To Dirs.ListCount - 1
       Debug.Print "parent " & Dirs.List(i)
       EnumChildNodes Node, Dirs.List(i), Dirs.List(i), Dirs.ListCount
    Next
End Sub

Private Sub EnumChildNodes(parent As Node, key As String, text As String, count As Integer)

If parent.Children < count Then
    Debug.Print key
    On Error GoTo err:
    GenView.Nodes.Add parent, tvwChild, key, Mid(text, InStrRev(text, "\") + 1, Len(text) - InStrRev(text, "\")), 2
    If CountSubDirs(key) > 0 Then 'add dummy node
        GenView.Nodes.Add key, tvwChild, , "dummy"
    End If
Exit Sub
err:
    MsgBox err.Description & " " & parent.key & " " & key
    Dim i As Integer
    For i = 1 To parent.Children
        If GenView.Nodes(i).key = key Then
            MsgBox "found " & i & " " & key
        End If
    Next
    Exit Sub
End If
End Sub

Private Sub GenView_NodeClick(ByVal Node As MSComctlLib.Node)
TempDir = Node.key
End Sub

Public Function CountSubDirs(path As String) As Integer
On Error GoTo Derr
SubDirs.path = path
CountSubDirs = SubDirs.ListCount
Exit Function
Derr:
SubDirs.path = "c:"
End Function
