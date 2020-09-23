VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "OCX/DLL Detect"
   ClientHeight    =   3375
   ClientLeft      =   135
   ClientTop       =   420
   ClientWidth     =   5535
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   5535
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton Command3 
      Caption         =   "&Quit"
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   2880
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   600
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   3960
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   120
      Top             =   3480
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&About"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Select a file"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'OCX/DLL Detect
'by Dennis "hax" Wollentin

Dim strContents As String
Dim x As Long
Dim y As Long
Dim z As String
Dim q As String
Dim u As Long
Dim strFilename As String
Dim L As Long

Sub Detect()
On Error Resume Next
x = 0
y = 0
z = ""
q = ""
u = 0

List1.Clear
Open strFilename For Binary As #1
strContents = Space(LOF(1))
Get #1, , strContents
Close #1

y = InStr(1, strContents, ".DLL")
u = y
5:
u = u - 1
q = Mid(strContents, u)
Text1.Text = q
If Len(Text1.Text) = 0 Then
x = u
GoTo 4
Else
GoTo 5
End If
4 z = Mid(strContents, u + 1)
List1.AddItem z

y = InStr(y + 4, strContents, ".DLL")
If y = 0 Then
GoTo detectocx
Else
u = y
q = ""
Text1.Text = ""
GoTo 5
End If

detectocx:
y = InStr(1, strContents, ".OCX")
u = y
6:
u = u - 1
q = Mid(strContents, u)
Text1.Text = q
If Len(Text1.Text) = 0 Then
x = u
GoTo 7
Else
GoTo 6
End If
7 z = Mid(strContents, u + 1)
List1.AddItem z

y = InStr(y + 4, strContents, ".OCX")
If y = 0 Then
GoTo detectdll
Else
u = y
q = ""
Text1.Text = ""
GoTo 6
End If

detectdll:
y = InStr(1, strContents, ".dll")
u = y
8:
u = u - 1
q = Mid(strContents, u)
Text1.Text = q
If Len(Text1.Text) = 0 Then
x = u
GoTo 11
Else
GoTo 8
End If
11 z = Mid(strContents, u + 1)
List1.AddItem z

y = InStr(y + 4, strContents, ".dll")
If y = 0 Then
GoTo detectsmallocx
Else
u = y
q = ""
Text1.Text = ""
GoTo 8
End If

detectsmallocx:
y = InStr(1, strContents, ".ocx")
u = y
25:
u = u - 1
q = Mid(strContents, u)
Text1.Text = q
If Len(Text1.Text) = 0 Then
x = u
GoTo 13
Else
GoTo 25
End If
13 z = Mid(strContents, u + 1)
List1.AddItem z

y = InStr(y + 4, strContents, ".ocx")
If y = 0 Then
If L = 1 Then
L = 0
Exit Sub
Else
L = 1
Detect
End If
Else
u = y
q = ""
Text1.Text = ""
GoTo 25
End If

End Sub

Private Sub Command1_Click()
List1.Clear
CommonDialog1.Filter = "Executable (*.exe)|*.exe|" & _
"All Files (*)|*|"
CommonDialog1.FilterIndex = 1
CommonDialog1.ShowOpen
If Len(CommonDialog1.filename) = 0 Then
Exit Sub
Else
strFilename = CommonDialog1.filename
End If
List1.Clear
Detect
End Sub

Private Sub Command2_Click()
Form2.Show
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

