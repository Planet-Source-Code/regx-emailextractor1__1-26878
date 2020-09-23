VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Email Extractor"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   9765
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Remove Selected"
      Height          =   315
      Left            =   1440
      TabIndex        =   9
      Top             =   5400
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Clear"
      Height          =   300
      Left            =   120
      TabIndex        =   8
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save(append) to file"
      Height          =   330
      Left            =   6960
      TabIndex        =   7
      Top             =   5400
      Width           =   2775
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4440
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Text(*.txt)|*.txt"
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   315
      Left            =   8640
      TabIndex        =   6
      Top             =   120
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BackColor       =   &H00404000&
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   9705
      TabIndex        =   5
      Top             =   5865
      Width           =   9765
   End
   Begin VB.ListBox List2 
      BackColor       =   &H00404000&
      ForeColor       =   &H00E0E0E0&
      Height          =   4740
      Left            =   6960
      TabIndex        =   4
      Top             =   600
      Width           =   2775
   End
   Begin VB.TextBox txtStartUrl 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "http://"
      Top             =   120
      Width           =   7575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   5160
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "Form1.frx":0000
      Top             =   5400
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go"
      Height          =   315
      Left            =   7800
      TabIndex        =   1
      Top             =   120
      Width           =   705
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00400000&
      ForeColor       =   &H00E0E0E0&
      Height          =   4740
      Left            =   120
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   600
      Width           =   6735
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   3720
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      RequestTimeout  =   5
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Needs reference to Microsoft VBscript Regular Expressions I recomend ver 5.5.
'

' There are better ways to do this, I will try to make suggestions in the comments
'
Dim baseurl As String ' var to store base url so we can build the full path
Dim dVisited As Dictionary ' Dictionary to hold visited urls
Dim dEmail As Dictionary ' dictionary to hold emails
' We are putting the emails in a list also, for user feed back
'It would be less momery intensive and faster to just keep these in the dictionry object
'which allows to easily tell if the email already exist
Dim regxPage ' var to hold regular expression to extract urls
Dim regxEmail ' var to hold regular expression to extract emails
Dim Match, Matches ' we use these to store are regx matches
' Regular expressions are super powerfull and have been a part of unix for a long time
' goto the form load event to see the regex initialization
'   to learn more about regular expressions and to download the latest scripting runtime see
'   http://msdn.microsoft.com/scripting/default.htm?/scripting/vbscript/techinfo/vbsdocs.htm
Dim email, pageurl As String
' The above is only because dictionary.exist does not work directly on Match var
Dim stopcrawl As Integer ' Used to exit crawl loop
Private Sub Command1_Click()
stopcrawl = 0 ' set stop crawl so we do not exit loop
If txtStartUrl.Text & "" = "" Then
    MsgBox "Please enter a Starting URL", vbOKOnly, "Oops"
    Exit Sub
ElseIf txtStartUrl.Text & "" = "http://" Then
MsgBox "Please enter a Starting URL", vbOKOnly, "Oops"
    Exit Sub
End If
' the above should really check for a valid url, but I am a lazy PERL programmer
List1.AddItem txtStartUrl 'add item to list
crawl ' and were off

End Sub
Sub crawl()
While List1.ListCount > 0 ' loop while list has data
    If stopcrawl = 1 Then GoTo exitcrawl ' is the user trying to stop the prog?
    getpage List1.List(0) ' This is the heart of the prog, except for the regx
                          ' stuff in the form load event
    List1.RemoveItem (0)  ' remove item from list
Wend
 MsgBox "Url list has been exhausted, try a startink page with more links"
exitcrawl:
End Sub
Sub getpage(page As String)
On Error Resume Next
' check if the current page has been visited
If dVisited.Exists(page) Then
    Picture1.Cls
    Picture1.Print "Skipping - Been there done that"
    'List1.RemoveItem (0)
    Exit Sub
Else
    dVisited.Add page, 1 ' add page to dVisited dictionary
    Picture1.Cls
    Picture1.Print "Cached=" & List1.ListCount & "  Visited=" & dVisited.Count & "  Emails=" & List2.ListCount
    Form1.Caption = page
End If

baseurl = getpath(page) ' build full url - see getpath
Text1 = ""
If List1.ListCount > 5000 Then Exit Sub ' sets the maximum cache (so we don't run out of mem)
Text1.Text = Inet1.OpenURL(page) ' get the html (no need to build path since list1 contains full paths)
' search for links
' Oh yes, after defining a regx pattern we just pass the text we want to search to
' the execute method and we get back a collection of matches.
' This collection for instance will contain all the urls in the text excluding
' targets such as www.someurl.com/somepage#targetname
' now you know why perl is so popular (very good regx support - Yes, better than this)
    Set Matches = regxPage.Execute(Text1.Text)    ' Execute search.
    For Each Match In Matches     ' Iterate Matches collection.
        pageurl = Match
        ' check if page is a full or relative url and take appropriate action
        ' the mid function here just removes the href=
        ' in perl we would do that in the regx directly, but
        ' VB uses a submatch collection when defining grouping in a regx
        ' wich is cumbersome and very lame
        ' It could still be done with submatches, but since we know what we want to remove
        ' and it never changes this way is better.
        If InStr(1, pageurl, "http://", vbTextCompare) Then
            If dVisited.Exists(pageurl) = False Then List1.AddItem Mid(pageurl, 7)
        Else
            If dVisited.Exists(baseurl & Mid(pageurl, 7)) = False Then List1.AddItem baseurl & Mid(pageurl, 7)
        End If
    Next
' search for email
 
    Set Matches = regxEmail.Execute(Text1.Text)    ' Execute search.
    For Each Match In Matches     ' Iterate Matches collection.
        ' check if email exist
         email = Match
Debug.Print email & dEmail.Exists(email)
        If dEmail.Exists(email) = False Then
            dEmail.Add (email), 1
            List2.AddItem email
            Beep
        End If
    Next
End Sub

Function getpath(url As String) As String
' look for the last / and get a string up to that location
lastbar = InStrRev(url, "/")
tmppath = Mid(url, 1, lastbar)
If tmppath = "http://" Then tmppath = url ' full path already so return url
getpath = tmppath
End Function


Private Sub Command2_Click()
stopcrawl = 1
End Sub

Private Sub Command3_Click()
CommonDialog1.InitDir = AppPath
CommonDialog1.ShowSave
If CommonDialog1.FileName & "" = "" Then
    MsgBox "You must enter a file name enter a filename", vbOKOnly, "Couldn't save file"
Else
 Close #1
    Open CommonDialog1.FileName For Append As #1
    For a = 0 To List2.ListCount
      Print #1, List2.List(a)
    Next a
    Close #1
    MsgBox "File Saved to " & CommonDialog1.FileName, vbOKOnly, "Done"
End If
End Sub

Private Sub Command4_Click()
List1.Clear
End Sub

Private Sub Command5_Click()
' remove selected items from list
' since the list indexes change every time we remove an item this is a little
' tricky, but not to bad
Dim a As Integer
For a = 0 To List1.ListCount - 1
recheck:
 If a > List1.ListCount - 1 Then Exit Sub
    If List1.Selected(a) = True Then
        List1.RemoveItem (a)
        GoTo recheck
    End If
Next a
End Sub

Private Sub Form_Load()
'initialize dictionary and regx objects
Set dVisited = CreateObject("Scripting.Dictionary")
dVisited.CompareMode = BinaryCompare
Set dEmail = CreateObject("Scripting.Dictionary")
dEmail.CompareMode = BinaryCompare
' define regular expresions
'regxPage pattern matches all href tags excluding targets
' for those of you new to regular expressions it basically says
' match where text=
' HREF=" followed by one or more characters that are not " or #
' The first part should be obvious except maybe for the "" wich is just an escaped double quote
' []square bracets are used to define a character class
' ^ is a negate operator so [^""#] means not " or #
' The last plus just means one or more of whatever it is trailing
' Take a look at the regxEmail pattern and see if you can tell what it is saying
' If you would like me to put together a better Regular Expressions tutorial let me know.
    Set regxPage = New RegExp            ' Create a regular expression.
    regxPage.Pattern = "HREF=""[^""#]+[.][^""#]+"      ' Set pattern."
    regxPage.IgnoreCase = True         ' Set case insensitivity.
    regxPage.Global = True           ' Set global applicability.
   
    Set regxEmail = New RegExp            ' Create a regular expression.
    regxEmail.Pattern = "[a-z0-9]+@[a-z0-9]+[.][.a-z0-9]+"      ' Set pattern."
    regxEmail.IgnoreCase = True         ' Set case insensitivity.
    regxEmail.Global = True           ' Set global applicability.
End Sub

Public Function AppPath() As String
AppPath = App.Path & IIf(Right(App.Path, 1) = "\", "", "\")
End Function
