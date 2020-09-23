VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "VB LCD"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4695
   BeginProperty Font 
      Name            =   "Fixedsys"
      Size            =   9
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
   ScaleHeight     =   3615
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Welcome to VB LCD"
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.CommandButton Command5 
         Caption         =   "Down"
         Height          =   255
         Left            =   3480
         TabIndex        =   9
         Top             =   2160
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Up"
         Height          =   255
         Left            =   3480
         TabIndex        =   8
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Initialize LCD"
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   720
         Width           =   1935
      End
      Begin VB.ListBox List1 
         Height          =   735
         ItemData        =   "frmMain.frx":0CCA
         Left            =   360
         List            =   "frmMain.frx":0CE0
         TabIndex        =   2
         Top             =   1680
         Width           =   3015
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Go!"
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   3000
         Width           =   1935
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Stop!"
         Height          =   375
         Left            =   2400
         TabIndex        =   4
         Top             =   3000
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Scroll LCD"
         Height          =   225
         Left            =   3440
         TabIndex        =   10
         Top             =   1920
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackColor       =   &H00808080&
         Caption         =   " Step one: Initialize LCD"
         Height          =   225
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   4455
      End
      Begin VB.Label Label2 
         BackColor       =   &H00808080&
         Caption         =   " Step two: Choose a Demo"
         Height          =   225
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   4440
      End
      Begin VB.Label Label3 
         BackColor       =   &H00808080&
         Caption         =   " Step three: Go"
         Height          =   225
         Left            =   120
         TabIndex        =   5
         Top             =   2640
         Width           =   4440
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''
''    frmMain.frm    ''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''
''  HD44780 LCD Example                          ''
''  By: Eddy Richard                             ''
'''''''''''''''''''''''''''''''''''''''''''''''''''
''  © 2002 Guy In Green Shirt, Inc.              ''
''  Check out http://www.guyingreenshirt.com     ''
''  This code is made Public Domain by me,       ''
''  so long as the previous text is left intact. ''
'''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Quit As Boolean
Dim Up As Boolean
Dim Down As Boolean
Private Sub Command1_Click()
LCD_Init
End Sub
Private Sub Command2_Click()
Dim aScreen As String
Dim Screen1 As String
Dim Screen2 As String
Dim Screen3 As String
Dim Screen4 As String
Dim Screen5 As String
Dim CurScreen As Integer
Dim LongLine As String
Dim I As Integer
Dim J As Integer
Select Case List1.ListIndex + 1
  Case 1
    'display sample title screen (with "wipe" effects)
    aScreen = "       VB_LCD       " & _
              "   Demonstration    " & _
              "      Program       " & _
              "    Version 0.40    "
    LCD_DisplayScreen aScreen
    Sleep 2000
    aScreen = " Copyright 2002 By  " & _
              "    Eddy Richard    " & _
              "                    " & _
              "       Enjoy!       "
    LCD_WipeOffLR
    LCD_WipeOnRL aScreen
    Sleep 2000
  Case 2
    'display sample "marquee"
    aScreen = "                    " & _
              "The top line of this" & _
              "screen should be    " & _
              "scrolling away ...  "
    LongLine = "                    This is a very long line of text, scrolling across the screen like a marquee.                    "
    I = 0
    LCD_DisplayScreen aScreen
    Command3.Visible = True
    Do Until Quit = True
      DoEvents
      I = I + 1
      LCD_DisplayRow 1, Mid(LongLine, I, 20)
      If I > Len(LongLine) - 20 Then I = 0
      Sleep 250
    Loop
    Quit = False
    Command3.Visible = False
  Case 3
    'display sample "flashing"
    Screen1 = "Use flashing text   " & _
              "to attract attention" & _
              "to a word or phrase." & _
              " Different Flashing "
    Screen2 = "Use          text   " & _
              "to attract attention" & _
              "to a word or phrase." & _
              "ÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿÿ"
    Command3.Visible = True
    Do Until Quit = True
      DoEvents
      LCD_DisplayScreen Screen1
      Sleep 250
      LCD_DisplayScreen Screen2
      Sleep 250
    Loop
    Quit = False
    Command3.Visible = False
  Case 4
    'display sample "bar graph"
    aScreen = "  Sample Bar Graph  " & _
              "                    " & _
              "                    " & _
              "LO                HI"
    Dim Bar(5) As Integer
    Bar(0) = Asc(" ")
    Bar(1) = LCD_CHAR_BAR1
    Bar(2) = LCD_CHAR_BAR2
    Bar(3) = LCD_CHAR_BAR3
    Bar(4) = LCD_CHAR_BAR4
    Bar(5) = LCD_CHAR_BAR5
    LCD_DisplayScreen aScreen
    Command3.Visible = True
    Do Until Quit = True
      DoEvents
      For I = 1 To 20
        For J = 1 To 5
          LCD_Cursor 3, I
          LCD_DisplayCharacter Chr(Bar(J))
          Sleep 10
        Next J
      Next I
      For I = 20 To 1 Step -1
        For J = 5 To 1 Step -1
          LCD_Cursor 3, I
          LCD_DisplayCharacter Chr(Bar(J - 1))
          Sleep 10
        Next J
      Next I
    Loop
    Quit = False
    Command3.Visible = False
  Case 5
    'display sample mult-page screen (with "arrows")
    Screen1 = "This is a sample    " & _
              "multi-page screen.  " & _
              "Note the arrow in   " & _
              "the corners, which " & Chr(6)
    Screen2 = "indicate that there" & Chr(5) & _
              "is more text avail- " & _
              "able to be viewed,  " & _
              "by pressing the    " & Chr(6)
    Screen3 = "arrow keys.  The   " & Chr(5) & _
              "arrows only show up " & _
              "when more screens   " & _
              "can be accessed.   " & Chr(6)
    Screen4 = "When the last and/ " & Chr(5) & _
              "or first screen is  " & _
              "reached, the arrow  " & _
              "disappears.         "
    CurScreen = 1
    Command3.Visible = True
    Command4.Visible = True
    Command5.Visible = True
    Label4.Visible = True
    Do Until Quit = True
      DoEvents
      Select Case CurScreen
         Case 1
           LCD_DisplayScreen Screen1
         Case 2
           LCD_DisplayScreen Screen2
         Case 3
           LCD_DisplayScreen Screen3
         Case 4
           LCD_DisplayScreen Screen4
      End Select
      If Quit = True Then
        Quit = False
        GoTo Hell
      ElseIf Down = True Then
        Down = False
        CurScreen = CurScreen + 1
        If CurScreen > 4 Then CurScreen = 4
      ElseIf Up = True Then
        Up = False
        CurScreen = CurScreen - 1
        If CurScreen < 1 Then CurScreen = 1
      End If
    Loop
  Case 6
    'display sample menu system
    Screen1 = "---- MAIN MENU -----" & _
              "~Option Number 1    " & _
              " Option Number 2    " & _
              " Option Number 3   " & Chr(6)
    Screen2 = "---- MAIN MENU -----" & _
              " Option Number 1    " & _
              "~Option Number 2    " & _
              " Option Number 3   " & Chr(6)
    Screen3 = "---- MAIN MENU -----" & _
              " Option Number 1    " & _
              " Option Number 2    " & _
              "~Option Number 3   " & Chr(6)
    Screen4 = "---- MAIN MENU -----" & _
              " Option Number 2   " & Chr(5) & _
              " Option Number 3    " & _
              "~Option Number 4   " & Chr(6)
    Screen5 = "---- MAIN MENU -----" & _
              " Option Number 3   " & Chr(5) & _
              " Option Number 4    " & _
              "~Option Number 5    "
    CurScreen = 1
    Command3.Visible = True
    Command4.Visible = True
    Command5.Visible = True
    Label4.Visible = True
    LCD_CursorOn
    Do Until Quit = True
      DoEvents
      Select Case CurScreen
        Case 1
          LCD_DisplayScreen Screen1
          LCD_Cursor 2, 1
        Case 2
          LCD_DisplayScreen Screen2
          LCD_Cursor 3, 1
        Case 3
          LCD_DisplayScreen Screen3
          LCD_Cursor 4, 1
        Case 4
          LCD_DisplayScreen Screen4
          LCD_Cursor 4, 1
        Case 5
          LCD_DisplayScreen Screen5
          LCD_Cursor 4, 1
      End Select
      If Quit = True Then
        Quit = False
        GoTo Hell
      ElseIf Down = True Then
        Down = False
        CurScreen = CurScreen + 1
        If CurScreen > 5 Then CurScreen = 5
      ElseIf Up = True Then
        Up = False
        CurScreen = CurScreen - 1
        If CurScreen < 1 Then CurScreen = 1
      End If
    Loop
    LCD_CursorOff
  Case Else
    MsgBox "Please select a demo", vbOKOnly Or vbCritical, "Select demo"
End Select
Hell:
Command3.Visible = False
Command4.Visible = False
Command5.Visible = False
Label4.Visible = False
End Sub
Private Sub Command3_Click()
Quit = True
End Sub
Private Sub Command4_Click()
Up = True
End Sub
Private Sub Command5_Click()
Down = True
End Sub
