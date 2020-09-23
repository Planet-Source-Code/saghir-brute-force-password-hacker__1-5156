VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Numbers"
   ClientHeight    =   3075
   ClientLeft      =   3195
   ClientTop       =   3195
   ClientWidth     =   5370
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   5370
   Begin VB.CommandButton Command6 
      Caption         =   "&Clear List"
      Height          =   375
      Left            =   2640
      TabIndex        =   40
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "E&xport to Main List"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   2520
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   3480
      TabIndex        =   1
      Top             =   480
      Width           =   1695
   End
   Begin TabDlg.SSTab s 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5325
      _ExtentX        =   9393
      _ExtentY        =   5318
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "&Random"
      TabPicture(0)   =   "Form2.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Command4"
      Tab(0).Control(1)=   "Text9"
      Tab(0).Control(2)=   "Option3"
      Tab(0).Control(3)=   "Option2"
      Tab(0).Control(4)=   "Option1"
      Tab(0).Control(5)=   "Text8"
      Tab(0).Control(6)=   "Label12"
      Tab(0).Control(7)=   "Label11"
      Tab(0).Control(8)=   "Label8"
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "&Squencial"
      TabPicture(1)   =   "Form2.frx":08E6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Text1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Text2"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Command2"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "&Input"
      TabPicture(2)   =   "Form2.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label4"
      Tab(2).Control(1)=   "Label5"
      Tab(2).Control(2)=   "Label6"
      Tab(2).Control(3)=   "Label7"
      Tab(2).Control(4)=   "Text3"
      Tab(2).Control(5)=   "Text4"
      Tab(2).Control(6)=   "Text5"
      Tab(2).Control(7)=   "Text6"
      Tab(2).Control(8)=   "Command3"
      Tab(2).Control(9)=   "List2"
      Tab(2).Control(10)=   "List3"
      Tab(2).Control(11)=   "List4"
      Tab(2).Control(12)=   "List5"
      Tab(2).ControlCount=   13
      TabCaption(3)   =   "&Combinations"
      TabPicture(3)   =   "Form2.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label9"
      Tab(3).Control(1)=   "Label10"
      Tab(3).Control(2)=   "Label13"
      Tab(3).Control(3)=   "Label14"
      Tab(3).Control(4)=   "Label15"
      Tab(3).Control(5)=   "Label16"
      Tab(3).Control(6)=   "Label17"
      Tab(3).Control(7)=   "Label18"
      Tab(3).Control(8)=   "Text7"
      Tab(3).Control(9)=   "Text10"
      Tab(3).Control(10)=   "Text11"
      Tab(3).Control(11)=   "Text12"
      Tab(3).Control(12)=   "Command5"
      Tab(3).ControlCount=   13
      Begin VB.ListBox List5 
         Height          =   1815
         Left            =   -71520
         TabIndex        =   44
         Top             =   480
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ListBox List4 
         Height          =   1815
         Left            =   -71520
         TabIndex        =   43
         Top             =   480
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ListBox List3 
         Height          =   1815
         Left            =   -71520
         TabIndex        =   42
         Top             =   480
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ListBox List2 
         Height          =   1815
         Left            =   -71520
         TabIndex        =   41
         Top             =   480
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Add"
         Height          =   375
         Left            =   -73440
         TabIndex        =   36
         Top             =   2520
         Width           =   975
      End
      Begin VB.TextBox Text12 
         Height          =   375
         Left            =   -72240
         TabIndex        =   31
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox Text11 
         Height          =   375
         Left            =   -73200
         TabIndex        =   30
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   -74160
         TabIndex        =   29
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   -74880
         TabIndex        =   28
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Add"
         Height          =   375
         Left            =   -73440
         TabIndex        =   26
         Top             =   2520
         Width           =   975
      End
      Begin VB.TextBox Text9 
         Height          =   375
         Left            =   -73440
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "1000"
         Top             =   1800
         Width           =   1335
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Eight"
         Height          =   255
         Left            =   -74760
         TabIndex        =   24
         Top             =   2280
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Six "
         Height          =   255
         Left            =   -74760
         TabIndex        =   23
         Top             =   1920
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Four"
         Height          =   255
         Left            =   -74760
         TabIndex        =   22
         Top             =   1560
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   -74760
         TabIndex        =   20
         Text            =   "1000"
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Add"
         Height          =   375
         Left            =   -73440
         TabIndex        =   13
         Top             =   2520
         Width           =   975
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   -73560
         TabIndex        =   12
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   -73560
         TabIndex        =   11
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   -73560
         TabIndex        =   10
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   -73560
         TabIndex        =   9
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Add"
         Height          =   375
         Left            =   1560
         TabIndex        =   5
         Top             =   2520
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         Text            =   "1000"
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1560
         TabIndex        =   3
         Text            =   "1"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label18 
         Caption         =   ")"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   -72600
         TabIndex        =   39
         Top             =   840
         Width           =   135
      End
      Begin VB.Label Label17 
         Caption         =   "("
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   -74160
         TabIndex        =   38
         Top             =   840
         Width           =   135
      End
      Begin VB.Label Label16 
         Caption         =   "Example:-  ABC-5930-XYZ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   37
         Top             =   1920
         Width           =   2895
      End
      Begin VB.Label Label15 
         Caption         =   "End  With"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   -72240
         TabIndex        =   35
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label14 
         Caption         =   "To"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   -72960
         TabIndex        =   34
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label13 
         Caption         =   "Numbers From"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   -74040
         TabIndex        =   33
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Start With"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   -74880
         TabIndex        =   32
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "Number and Alphabet Combinations"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   27
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label Label12 
         Caption         =   "Select Digits:-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   21
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Recammended"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   -73200
         TabIndex        =   19
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "How many Numbers are recuired?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   18
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label Label7 
         Caption         =   "Input 4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74400
         TabIndex        =   17
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Input 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74400
         TabIndex        =   16
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Input 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74400
         TabIndex        =   15
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Input 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74400
         TabIndex        =   14
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Set Limits:-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "To:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Start from:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   1200
         Width           =   975
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'I wanted to write a password recovery but i ended up with this
'file recovery program.What i mean by that is I can only
'open ,decrypt ,unprotect a file or program But i can,t get the read password.
Private Sub Command1_Click()
MousePointer = 11
Dim x As Integer
x = 0
If List1.ListCount = 0 Then
MousePointer = 0
MsgBox "There are no numbers in The list"
Exit Sub
Else
Do Until x = List1.ListCount
Form1.List1.AddItem Form2.List1.List(x)
x = x + 1
Loop
End If
x = 0
MousePointer = 0
Unload Me
End Sub

Private Sub Command2_Click()
'sequential numbers
Dim f, last As Variant
If f = Empty Then f = Text1.Text
Do Until f = Text2.Text
f = f + 1
If List1.ListCount >= 32000 Then
'I added four list boxes  to be able to have 100000 pass's at a time
'but I'm not going to use all of them because it will probally hang windows
'just tried it to see if i could
GoTo Next1
'Form2.MousePointer = 0
'MsgBox "This password hacker allows 32000 passwords at a time as it wasn't desiened to hack into pantagon"
End If
List1.AddItem f
Form2.MousePointer = 11
Loop
Next1:
Do Until f = Text2.Text
If List2.ListCount > 32000 Then GoTo Next2
List2.AddItem x
x = x + 1
Loop
Next2:
Do Until f = Text2.Text
If List3.ListCount > 32000 Then GoTo next3
List3.AddItem x
x = x + 1
Loop
next3:
Do Until f = Text2.Text
If List4.ListCount > 32000 Then GoTo next4
List4.AddItem x
x = x + 1
Loop
next4:
Do Until f = Text2.Text
If List5.ListCount > 32000 Then GoTo next5
List5.AddItem x
x = x + 1
Loop
next5:
Form2.MousePointer = 0
End Sub

Private Sub Command3_Click()
If List1.ListCount >= 32000 Then
Form2.MousePointer = 0
MsgBox "This password hacker allows 32000 passwords at a time as it wasn't desiened to hack into pantagon"
Exit Sub
End If
'if statement's used to avoid blank lines
If Text3.Text = "" Then
Else
List1.AddItem Text3.Text
End If
If Text4.Text = "" Then
Else
List1.AddItem Text4.Text
End If
If Text5.Text = "" Then
Else
List1.AddItem Text5.Text
End If
If Text6.Text = "" Then
Else
List1.AddItem Text6.Text
End If
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""

End Sub


Private Sub Command4_Click()
Static x As Integer
If x = Empty Then x = 0
   
    Dim ranNum As Variant
        Dim ranN As Integer
        Dim ran  As Integer
        Dim ran2 As Integer
'this means creat passwords untill it reaches the number in text8
Do Until x = Text8.Text
   Form2.MousePointer = 11
      Randomize
      If List1.ListCount >= 32000 Then
Form2.MousePointer = 0
MsgBox "This file Recovery tool allows 32000 passwords at a time as it wasn't desiened to hack into pantagon"
Exit Sub
End If
   'four digit numbers
   If Text9.Text = 1000 Then
   ranN = (10000 * Rnd) + 1
   List1.AddItem ranN
   x = x + 1
   End If
   'six digit numbers
   If Text9.Text = 100000 Then
   ranN = (1000 * Rnd) + 1
   ran = (1000 * Rnd) + 1
ranNum = ranN & ran
   List1.AddItem ranNum
   x = x + 1
   End If
    'Eight digit numbers
   If Text9.Text = 10000000 Then
   ranN = (1000 * Rnd) + 1
   ran = (1000 * Rnd) + 1
      ran2 = (100 * Rnd) + 1

ranNum = ranN & ran & ran2
   List1.AddItem ranNum
   x = x + 1
   End If
   Loop
   Form2.MousePointer = 0
   x = 0
End Sub

Private Sub Command5_Click()
Dim start, finish As Variant
If start = Empty Then start = Text10.Text
Do Until start = Text11.Text
start = start + 1
If List1.ListCount >= 32000 Then
Form2.MousePointer = 0
MsgBox "This password hacker allows 32000 passwords at a time as it wasn't desiened to hack into pantagon"
Exit Sub
End If
List1.AddItem Text7.Text & start & Text12.Text
Form2.MousePointer = 11
Loop
Form2.MousePointer = 0
End Sub

Private Sub Command6_Click()
List1.Clear
End Sub

Private Sub Option1_Click()
Text9.Text = 1000
End Sub
Private Sub Option2_Click()
Text9.Text = 100000
End Sub

Private Sub Option3_Click()
Text9.Text = 10000000
End Sub

Private Sub Text8_Change()
If Text8.Text = 1000 Then
Label11.Visible = True
Else
Label11.Visible = False
End If
End Sub
