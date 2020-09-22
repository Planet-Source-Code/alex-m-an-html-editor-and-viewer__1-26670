VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Web Viewer v2.0"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11310
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   11310
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Forward"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   855
   End
   Begin SHDocVwCtl.WebBrowser WebView 
      Height          =   6375
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   11295
      ExtentX         =   19923
      ExtentY         =   11245
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click() 'start sub
On Error GoTo errorhandler 'If any errors, goto errorhandler and exit sub
WebView.GoBack 'tell website viewer to go "Back"
errorhandler: 'errorhandler (explained above)
End Sub ' exit sub

Private Sub Command2_Click() 'start sub
On Error GoTo errorhandler 'If any errors, goto errorhandler and exit sub
WebView.GoForward 'tell website viewer to go "Forward"
errorhandler: 'errorhandler (explained above)
End Sub ' exit sub

Private Sub Form_Load() 'start sub
WebView.Navigate App.Path & "\Temp.html" 'tells the website viewer to
                                         ' view the file saved before
                                         'when you click on the view it
                                         'menu item.
End Sub ' exit sub


