VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password Recovery by Saghir"
   ClientHeight    =   5295
   ClientLeft      =   3000
   ClientTop       =   1860
   ClientWidth     =   5790
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   5790
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   0
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   3840
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000008&
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   5775
      Begin VB.CommandButton Command11 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   4320
         TabIndex        =   35
         Top             =   4560
         Width           =   1335
      End
      Begin VB.CommandButton Command10 
         Caption         =   "&Clear List"
         Height          =   375
         Left            =   4320
         TabIndex        =   32
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox Text17 
         Height          =   375
         Left            =   5040
         TabIndex        =   31
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox Text16 
         Height          =   375
         Left            =   4440
         TabIndex        =   30
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox Text15 
         Height          =   375
         Left            =   3840
         TabIndex        =   29
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox Text14 
         Height          =   375
         Left            =   3240
         TabIndex        =   28
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox Text13 
         Height          =   375
         Left            =   2640
         TabIndex        =   27
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox Text12 
         Height          =   375
         Left            =   2040
         TabIndex        =   26
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox Text11 
         Height          =   375
         Left            =   1440
         TabIndex        =   25
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   840
         TabIndex        =   24
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   1920
         TabIndex        =   23
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox Text9 
         Height          =   375
         Left            =   5040
         TabIndex        =   22
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   4440
         TabIndex        =   21
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   3840
         TabIndex        =   20
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   3240
         TabIndex        =   19
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   2640
         TabIndex        =   18
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2040
         TabIndex        =   17
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Clear list2"
         Height          =   375
         Left            =   480
         TabIndex        =   15
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Pause"
         Height          =   375
         Left            =   4320
         TabIndex        =   14
         Top             =   3840
         Width           =   1335
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&Numbers"
         Height          =   375
         Left            =   4320
         TabIndex        =   13
         Top             =   3480
         Width           =   1335
      End
      Begin VB.CommandButton Command6 
         Caption         =   "A&bout"
         Height          =   375
         Left            =   4320
         TabIndex        =   12
         Top             =   4200
         Width           =   1335
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Text            =   "3"
         Top             =   1440
         Width           =   255
      End
      Begin VB.CommandButton Command3 
         Caption         =   "File to Recover"
         Height          =   375
         Left            =   4320
         TabIndex        =   9
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Recover"
         Height          =   375
         Left            =   4320
         TabIndex        =   7
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox T3 
         Height          =   375
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   2280
         Width           =   2295
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Load from file"
         Height          =   375
         Left            =   4320
         TabIndex        =   5
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1920
         TabIndex        =   4
         Top             =   1800
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Add to list"
         Height          =   375
         Left            =   4320
         TabIndex        =   3
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Save to &file"
         Height          =   375
         Left            =   4320
         TabIndex        =   2
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   0
         Top             =   3360
      End
      Begin VB.ListBox List1 
         Height          =   2205
         ItemData        =   "Form1.frx":08CA
         Left            =   1920
         List            =   "Form1.frx":08CC
         TabIndex        =   1
         Top             =   2760
         Width           =   2295
      End
      Begin VB.ListBox List2 
         Height          =   2205
         Left            =   120
         TabIndex        =   16
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000006&
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   4320
         TabIndex        =   34
         Top             =   5040
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000006&
         Caption         =   "No of passwords Loaded:"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   2280
         TabIndex        =   33
         Top             =   5040
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000008&
         Caption         =   "Select keys stokes And Put T3 (Capital T) in the place of the password.."
         ForeColor       =   &H8000000E&
         Height          =   615
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000006&
         Caption         =   "Enter More words"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   1920
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim key1, key2, key3, key4, key5 As String
Dim key6, key7, key8, key9, key10, key11, key12, key13, key14 As String
Private Sub Command1_Click()
'if statement is used so that blank lines don't get added to list
If Text1.Text = "" Then
Else
List1.AddItem Text1.Text
Text1.Text = ""
End If
listcounting
End Sub
Private Sub Command10_Click()
List1.Clear
listcounting
End Sub

Private Sub Command11_Click()
End
End Sub

Private Sub Command2_Click()
If List1.ListCount = 0 Then Exit Sub
Static count As Integer
Dim path As String
If count = Empty Then count = 0
path = App.path & "\" & count & ".txt"
'if there is a error then that the file don't
'already exist's ,so goto savefile
On Error GoTo savefile
Do
ChDir path
count = count + 1
Loop
Exit Sub
savefile:
Open path For Output As #1
'saving the list to a file
For X = 0 To List1.ListCount - 1
    Print #1, List1.List(X)
 Next
 Close
count = count + 1
Exit Sub
End Sub
Private Sub Command3_Click()
'this is used in case some one want's to
'recover a forgotten password of a zip file
'winzip only gives you one try if the file is a
'win zip self extractor and 2 or 3 if it is ".zip"
'if there is a file name in text3 then this program
'loads it after every try of file recovery
MsgBox "Only choose the file if the file closes it self after one try i.e. winzip pass protected zip, in all other cases run the program and put the keys that take you to the password screen,other wise leave this box empty"
With CommonDialog1
.Filter = "Zip files(*.exe)|*.exe"
.ShowOpen
 End With
Text3.Text = CommonDialog1.filename
End Sub
Private Sub Command4_Click()
Static delays As Integer
If delays = Empty Then delays = 0
On Error Resume Next
If Timer1.Enabled = False Then
GoTo r
Exit Sub
End If
  ', vbHide
T3.Text = List1.List(0)
T3.Refresh
On Error Resume Next
List2.AddItem List1.List(0)
List1.RemoveItem 0
List1.Refresh
delays = delays + 1
'give windows 3 seconds after trying 500 passwords
If delays = 500 Then
Command8_Click
Timer2.Enabled = True
delays = 0
Exit Sub
End If
hacksub
Exit Sub
r:
Timer2.Enabled = True
End Sub
Private Sub Command5_Click()
'I used CommonDialog because it's every one is
'used to it ,also it's  very easy to use it when coding
With CommonDialog2
.Filter = "Any Word List|*.*"
.ShowOpen
End With
If CommonDialog2.filename = "" Then
Exit Sub
Else
'change mouse pointer to hourglass
Form2.MousePointer = 11
'load words to list
Open CommonDialog2.filename For Input As 1
Do Until EOF(1)
    Line Input #1, st
    List1.AddItem st
Loop
Close 1
CommonDialog2.filename = ""
End If
Form2.MousePointer = 0
listcounting
End Sub

Private Sub Command6_Click()
Dim MsgBox1 As String
MsgBox1 = MsgBox("Coded by saghir Dec 1999", vbInformation + vbOKOnly, "About")
End Sub
Private Sub Command7_Click()
'show numbers form
Form2.Show
End Sub
Private Sub Command8_Click()
Timer1.Enabled = False
End Sub
Private Sub Command9_Click()
List2.Clear
End Sub
Private Sub Form_Load()
listcounting
Dim delay As Integer
Dim del As Boolean
End Sub
Private Sub getstringvalue()
key1 = Text2.Text
key2 = Text5.Text
key3 = Text6.Text
key4 = Text7.Text
key5 = Text8.Text
key6 = Text9.Text
key7 = Text10.Text
key8 = Text11.Text
key9 = Text12.Text
key10 = Text13.Text
key11 = Text14.Text
key12 = Text15.Text
key13 = Text16.Text
key14 = Text17.Text
End Sub
Private Sub Form_Terminate()
Unload Me
End Sub
Private Sub Text10_Change()
getstringvalue
End Sub
Private Sub Text11_Change()
getstringvalue
End Sub
Private Sub Text12_Change()
getstringvalue
End Sub
Private Sub Text13_Change()
getstringvalue
End Sub
Private Sub Text14_Change()
getstringvalue
End Sub
Private Sub Text15_Change()
getstringvalue
End Sub
Private Sub Text16_Change()
getstringvalue
End Sub
Private Sub Text17_Change()
getstringvalue
End Sub
Private Sub Text2_Change()
getstringvalue
End Sub
Private Sub Text5_Change()
getstringvalue
End Sub
Private Sub Text6_Change()
getstringvalue
End Sub
Private Sub Text7_Change()
getstringvalue
End Sub
Private Sub Text8_Change()
getstringvalue
End Sub
Private Sub Text9_Change()
getstringvalue
End Sub
Private Sub Timer1_Timer()
'repeat
Command4_Click
listcounting
'stop if no more passwords left
If T3.Text = "" Then Timer1.Enabled = False
End Sub
Private Sub Timer2_Timer()
If Text4.Text = 0 Then
Text4.Text = 3
Text4.Refresh
GoTo r
End If
r:
delay = Text4.Text
delay = delay - 1
Text4.Text = delay
Text4.Refresh
If delay < 1 Then
Timer2.Enabled = False
Timer1.Enabled = True
End If
End Sub
Private Sub hacksub()
If key1 = "T3" Then
key1 = ""
SendKeys T3
Else
SendKeys key1
End If
If key2 = "T3" Then
key2 = ""
SendKeys T3
Else
SendKeys key2
End If
If key3 = "T3" Then
key3 = ""
SendKeys T3
Else
SendKeys key3
End If
If key4 = "T3" Then
key4 = ""
SendKeys T3
Else
SendKeys key4
End If
If key5 = "T3" Then
key5 = ""
SendKeys T3
Else
SendKeys key5
End If
If key6 = "T3" Then
key6 = ""
SendKeys T3
Else
SendKeys key6
End If
If key7 = "T3" Then
key7 = ""
SendKeys T3
Else
SendKeys key7
End If
If key8 = "T3" Then
key8 = ""
SendKeys T3
Else
SendKeys key8
End If
If key9 = "T3" Then
key9 = ""
SendKeys T3
Else
SendKeys key9
End If
If key10 = "T3" Then
key10 = ""
SendKeys T3
Else
SendKeys key10
End If
If key11 = "T3" Then
key11 = ""
SendKeys T3
Else
SendKeys key11
End If
If key12 = "T3" Then
key12 = ""
SendKeys T3
Else
SendKeys key12
End If
If key13 = "T3" Then
key13 = ""
SendKeys T3
Else
SendKeys key13
End If
If key14 = "T3" Then
key14 = ""
SendKeys T3
Else
SendKeys key14
End If
getstringvalue
Shell CommonDialog1.filename
End Sub
Private Sub listcounting()
Label4.Caption = List1.ListCount
Label4.Refresh
End Sub

