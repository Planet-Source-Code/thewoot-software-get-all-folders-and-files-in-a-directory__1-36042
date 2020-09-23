VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Get Files and/or Folders"
   ClientHeight    =   2310
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   4125
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   4125
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      Height          =   1455
      Left            =   0
      TabIndex        =   7
      Top             =   4080
      Width           =   4095
   End
   Begin VB.DirListBox Dir1 
      Height          =   1440
      Left            =   0
      TabIndex        =   6
      Top             =   2520
      Width           =   4095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Main Menu"
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Get files and folders"
         Height          =   255
         Left            =   2160
         TabIndex        =   5
         Top             =   120
         Width           =   1815
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   1005
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   3855
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Text            =   "C:\"
         Top             =   480
         Width           =   3855
      End
      Begin VB.Label Label2 
         Caption         =   "Output:"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         Top             =   960
         Width           =   1695
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         DrawMode        =   16  'Merge Pen
         X1              =   120
         X2              =   3960
         Y1              =   855
         Y2              =   855
      End
      Begin VB.Line Line1 
         BorderStyle     =   6  'Inside Solid
         DrawMode        =   2  'Blackness
         X1              =   120
         X2              =   3960
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Directory:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Menu mnuStart 
      Caption         =   "&Begin Scan"
   End
   Begin VB.Menu mnuExit 
      Caption         =   "&Unload Example"
      NegotiatePosition=   3  'Right
      WindowList      =   -1  'True
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'// Created 06/19/2002
'// Gets all folders (and files) in a given directory
'// You may do as you wish to this code. I support the open-source code movement.
'// This might not be the fastest way to get all the folders/files, but it does its job, and is decent at what it does.
'// I wrote it just to see if I could do it. Plus, I have never found a good, easy-to-use sub like this one.
'// I hope it makes your work a little easier. Enjoy it.

Private Sub Form_Unload(Cancel As Integer)
'// Sometimes the form just hides when you click the X, so we don't like that ;]
Cancel = 1
MsgBox "Instead of clicking here, please click the menu 'Unload Example'", vbInformation, "Exit"
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuStart_Click()
'// If you don't want to clear the list, just remove the following line:
List1.Clear

'// If the check box has a check in it, then get the folders and files, if not, just get the folders
If Check1.Value = Checked Then
    Call Peek(Text1.Text, True, List1, Dir1, File1)
Else
    Call Peek(Text1.Text, False, List1, Dir1, File1)
End If

End Sub
Public Sub Peek(Location As String, GetFiles As Boolean, Listbox As Listbox, Dir As DirListBox, _
FileList As FileListBox)

'// This is the sub that you probably are interested in. I will do my best to explain it
'// You will need a path to 'peek' inside, a listbox, directory listbox, and file listbox
'// Usage:
'//     To get all files, and folders inside a directory: Call Peek("C:\Example",True,List1,Dir1,File1)
'//     To get only the folders inside a directory: Call Peek("C:\Example",False,List1,Dir1,File1)
'// To add this sub in your program, copy the Public Sub(Location....) to End sub, and paste in either a module,
'// or in a form, like this.

Dim intX As Integer, intW As Integer, intZ As Integer
Dim colFolders As New Collection

'// Remove these two following lines if you would like to have your DirListbox and FileListbox visible
Dir.Visible = False
FileList.Visible = False

'// Let's make sure that we have the same directory format...always a good idea ;]
If Right(Location, 1) <> "\" Then Location = Location & "\"

'// We need to have the directory and the file boxes with current paths
Dir.Path = Location
FileList.Path = Dir.Path

'// Add the path to the listbox, one of the purposes of this sub
'// DoEvents because I don't want to freeze the program up.
Listbox.AddItem Dir.Path
DoEvents

'// If the programer wants to get folders and files, then bigolly get the files!
'// Circle through a loop adding the filenames and paths of current directory to the listbox
If GetFiles = True Then
        FileList.Path = Dir.Path
        For intZ = 0 To FileList.ListCount - 1
                If Len(FileList.Path) <= 3 Then
                    Listbox.AddItem Dir.Path & FileList.List(intZ)
                Else
                    Listbox.AddItem Dir.Path & "\" & FileList.List(intZ)
                End If
        Next intZ
End If

'// Go through the current directory list, and add the paths to a collection (for later refrence)
For intX = 0 To Dir.ListCount - 1
    colFolders.Add Dir.List(intX)
Next intX

'// Circle through the collection (above)
'// If more directories exist, run this sub again
For intW = 1 To colFolders.Count
    Dir.Path = colFolders.Item(intW)
    Peek Dir.Path, GetFiles, Listbox, Dir, FileList
Next intW

'// (For memory saving purposes)
Set colFolders = Nothing
    
End Sub
