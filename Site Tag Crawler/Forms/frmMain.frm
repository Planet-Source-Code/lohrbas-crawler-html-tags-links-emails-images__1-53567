VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Crawler"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   5505
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox cImage 
      Caption         =   "Image"
      Height          =   255
      Left            =   4680
      TabIndex        =   9
      Top             =   2880
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CheckBox cEmails 
      Caption         =   "Email"
      Height          =   255
      Left            =   4680
      TabIndex        =   8
      Top             =   3240
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CheckBox cLinks 
      Caption         =   "Links"
      Height          =   255
      Left            =   4680
      TabIndex        =   7
      Top             =   2520
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Help"
      Height          =   255
      Left            =   4680
      TabIndex        =   6
      Top             =   2040
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4560
      Top             =   3480
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3240
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   255
      Left            =   4680
      TabIndex        =   3
      Top             =   840
      Width           =   735
   End
   Begin MSComctlLib.TreeView ListView 
      Height          =   2535
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   4471
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton bCrawl 
      Caption         =   "Crawl"
      Default         =   -1  'True
      Height          =   255
      Left            =   4680
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtLink 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "http://www.google.com/"
      Top             =   120
      Width           =   4455
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   4920
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Idle"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3600
      Width           =   4455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Threshold As Long

Private Sub bCrawl_Click()
    ' I wanted to add a threshold so that the program
    ' would repeatedly search, but it was too much
    ' trouble so I decided not to... If anyone
    ' figures out a recursive (or any efficient method)
    ' method to have a threshold on searching links,
    ' please contact me. (Contact information on
    ' (Tags) Module)
    
    ' Yes, yes, I know that the program still ends up
    ' searching for links/emails if either of those
    ' are checked and that it is a waste of time...
    
    ' However, this is just a sample project
    ' to go with the module to give you an
    ' example of what it can do. If your needs
    ' include saving a list/loading a list,
    ' and such sort of extra things... then
    ' make your own project with those needs met.
    
    On Error Resume Next
    
    Call cmdClear_Click ' Make sure it's clear to not error out
    
    If cImage.Value = 0 And cLinks.Value = 0 And cEmails.Value = 0 Then
        MsgBox "Please check areas to search for on the webpage.", vbOKOnly + vbExclamation, "Search"
        Exit Sub
    End If
    
    Timer1.Enabled = True
    Dim Links() As String, Img() As String, Link As String
    
    If cImage.Value = 1 Or cLinks.Value = 1 Then
    
    Links() = SearchTag("<a", txtLink.Text)
    
    Dim i As Long
    For i = 0 To UBound(Links())
        Link = ViewProperty(Links(0), Links(i), "href")
        If LCase(Left(Link, Len("mailto:"))) = LCase("mailto:") And cImage.Value = 1 Then
            Link = Mid(Link, Len("mailto:") + 1)
            ListView.Nodes.Add "emails", tvwChild, "Email" & i, Link
        Else
            If cLinks.Value = 0 Then GoTo Nexti
            ListView.Nodes.Add "links", tvwChild, "Link" & i, Link
        End If
Nexti:
    Next i
    
    End If
    
    Link = ""
    Img() = SearchTag("<img", txtLink.Text)
    
    For i = 0 To UBound(Img())
        'MsgBox Images(i)
        Link = ViewProperty(Img(0), Img(i), "src")
        If Not Trim(Link) = vbNullString Then ' No link
            ListView.Nodes.Add "images", tvwChild, "Image" & i, Link ' There is a link so images are added
        End If
    Next i
End Sub

Private Sub cmdClear_Click()
    ListView.Nodes.Clear
    Form_Load
End Sub

Private Sub Command1_Click()
    MsgBox "1) Type in the website address" & vbCrLf & _
    "2) Check the box next to the items that you want to search for" & vbCrLf & _
    "3) Click on Crawl" & vbCrLf & vbCrLf & _
    "This is a sample project and thus I didn't bother adding a save/load feature." & _
    " It is also missing a 'threshold' feature which I wanted to add but I didn't have" & _
    " time to work on it.", vbOKOnly + vbInformation, "Help"
End Sub

Private Sub Form_Load()
    ' Set up ListView basic format
    
    With ListView.Nodes
        .Add , , "crawler", "Crawler"
        
        .Add "crawler", tvwChild, "links", "Links"
        .Add "links", tvwNext, "images", "Images"
        .Add "images", tvwNext, "emails", "E-mail Addresses"
    
        .Item("crawler").Expanded = True
    End With
End Sub

Private Sub Timer1_Timer()
' Updates ProgressBar/Status label
' with current status
On Error Resume Next

Dim mx As Long, mn As Long

mx = Split(Status, Chr(1))(1)
mn = Split(Status, Chr(1))(0)

pb.Max = mx
pb.Min = mn

lblStatus.Caption = mn & "/" & mx & " checked."
End Sub
