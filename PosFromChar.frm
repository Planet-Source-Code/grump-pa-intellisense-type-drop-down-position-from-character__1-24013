VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   840
      ItemData        =   "PosFromChar.frx":0000
      Left            =   2040
      List            =   "PosFromChar.frx":0019
      TabIndex        =   2
      Top             =   2520
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   2640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   1125
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Type some stuff then press the ""."" period key to bring up the pseudo function list"
      Height          =   495
      Left            =   1080
      TabIndex        =   3
      Top             =   1560
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Const EM_POSFROMCHAR = &HD6

Function LoWord(ByVal dw As Long) As Integer
    If dw And &H8000& Then
        LoWord = dw Or &HFFFF0000
    Else
        LoWord = dw And &HFFFF&
    End If
End Function

Function HiWord(ByVal dw As Long) As Integer
    HiWord = (dw And &HFFFF0000) \ 65536
End Function

Private Sub Form_Load()
    Dim p As Long
    p = SendMessage(Text1.hwnd, EM_POSFROMCHAR, ByVal Text1.SelLength + Text1.SelStart, 0)
    Debug.Print "x: " & LoWord(p)
    Debug.Print "y: " & HiWord(p)
End Sub




Private Sub List1_Click()
'List1.ListIndex = List1.ListIndex + 1

End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Text1.SetFocus
Else
 If KeyCode = 13 Then
  Text1.Text = Text1.Text & List1.List(List1.ListIndex)
  List1.Visible = False
  
  
  
 End If
End If
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 190 And Shift = 0 Then
    Dim p As Long
    Text2.Text = Text1.Text + "I"
    p = SendMessage(Text2.hwnd, EM_POSFROMCHAR, ByVal Text1.SelStart, 0)

    List1.Left = LoWord(p) * 15 + 15 + Text1.Left
    List1.Top = HiWord(p) + 580
List1.Visible = True
Else
If KeyCode = 40 Then
List1.SetFocus

Else
List1.Visible = False
End If
End If

End Sub
