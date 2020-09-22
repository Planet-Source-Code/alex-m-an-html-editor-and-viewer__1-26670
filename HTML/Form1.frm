VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HTML Deluxe"
   ClientHeight    =   3615
   ClientLeft      =   3480
   ClientTop       =   3135
   ClientWidth     =   6015
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   6015
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3720
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   3615
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "Form1.frx":030A
      Top             =   0
      Width           =   6015
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu new 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu aa 
         Caption         =   "-"
      End
      Begin VB.Menu open 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu save 
         Caption         =   "&Save As"
         Shortcut        =   ^S
      End
      Begin VB.Menu ab 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu edit 
      Caption         =   "&Edit"
      Begin VB.Menu editt 
         Caption         =   "Cu&t"
         Index           =   0
      End
      Begin VB.Menu editt 
         Caption         =   "&Copy"
         Index           =   1
      End
      Begin VB.Menu editt 
         Caption         =   "&Paste"
         Index           =   2
      End
   End
   Begin VB.Menu qweb 
      Caption         =   "&Website"
      Begin VB.Menu view 
         Caption         =   "&View Website"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu help 
      Caption         =   "&Help"
      Begin VB.Menu helpme 
         Caption         =   "&Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu ac 
         Caption         =   "-"
      End
      Begin VB.Menu about 
         Caption         =   "&About"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub about_Click() 'start sub
frmAbout.Show 'show the about form
End Sub ' exit sub

Private Sub editt_Click(Index As Integer) 'start sub
    Select Case Index
        Case 0 'Cut
            Clipboard.Clear 'Clear the clipboard
            Clipboard.SetText Text1.SelText 'cut text from textbox->clipboard
            Text1.SelText = "" 'text box loses text that you cut from it
        Case 1 'Copy
            Clipboard.Clear 'Clear the clipboard
            Clipboard.SetText Text1.SelText 'copy text from textbox->clipboard
        Case 2 'Paste
            Text1.SelText = Clipboard.GetText() 'take text from clipboard and place it on the textbox
    End Select

End Sub ' exit sub

Private Sub exit_Click() 'start sub
End 'exit program
End Sub 'exit sub

Private Sub helpme_Click() 'start sub
Form3.Show 'show form3
End Sub ' exit sub

Private Sub new_Click() 'start sub
Text1.Text = "" 'clear text box
End Sub ' exit sub

Private Sub open_Click() 'start sub
Wrap$ = Chr$(13) + Chr$(10) 'Prepare to wrap characters
    CommonDialog1.Filter = "Website Files (*.HTML)|*.HTML" 'Filter everything but this type of file
    CommonDialog1.ShowOpen 'Show the "Open Dialog"
    If CommonDialog1.FileName <> "" Then 'If name is anything but "" (nothing)
        Form1.MousePointer = 11 'mouse pointer is Hourglass
        Open CommonDialog1.FileName For Input As #1 'The the selected file is opened
        On Error GoTo TooBig: 'any errors it will skip to "TooBig" below
        Do Until EOF(1) 'self explanatory
            Line Input #1, LineOfText$ 'reads file's text
            AllText$ = AllText$ & LineOfText$ & Wrap$ 'wrap all
        Loop 'loop
        
        Text1.Text = AllText$ 'puts text from file to textbox1
        Text1.Enabled = True 'textbox is now enabled
        
        
CleanUp: 'clean up
        Form1.MousePointer = 0 'bring cursor back to default
        Close #1 'close #1 (from above)
    End If
    Exit Sub
TooBig: 'Too Big
    MsgBox ("The specified file is to big!!!") 'Tell user file is too big
    Resume CleanUp: 'pretty much goto clean up
End Sub ' exit sub

Private Sub save_Click() 'start sub
    CommonDialog1.Filter = "Website Files (*.HTML)|*.HTML" 'Filter everything but this type of file
    CommonDialog1.ShowSave 'Show the "Save Dialog"
    If CommonDialog1.FileName <> "" Then 'Make sure filename is not ""
        Open CommonDialog1.FileName For Output As #1 'Save using output
        'save text as string
        Print #1, Text1.Text 'print text into file
        Close #1 'close #1
    End If
End Sub ' exit sub
Private Sub view_Click() 'start sub
Open "Temp.html" For Output As #1 'open a file for output
Print #1, Text1.Text 'print text (HTML Code) into it
Close #1 'closes that file
Form2.Show 'show form2
End Sub ' exit sub
