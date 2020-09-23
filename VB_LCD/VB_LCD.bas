Attribute VB_Name = "modVB_LCD"
''''''''''''''''''''''''
''   modVB_LCD.bas   ''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''
''  HD44780 LCD Example                          ''
''  By: Eddy Richard                             ''
'''''''''''''''''''''''''''''''''''''''''''''''''''
''  © 2002 Guy In Green Shirt, Inc.              ''
''  Check out http://www.guyingreenshirt.com     ''
''  This code is made Public Domain by me,       ''
''  so long as the previous text is left intact. ''
'''''''''''''''''''''''''''''''''''''''''''''''''''
Declare Sub vbOut Lib "WIN95IO.DLL" (ByVal nPort As Integer, ByVal nData As Integer)
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Type LCD_CHAR
  ROW_0 As Long
  ROW_1 As Long
  ROW_2 As Long
  ROW_3 As Long
  ROW_4 As Long
  ROW_5 As Long
  ROW_6 As Long
  ROW_7 As Long
End Type
Public Type LCD_CUSTOM_i
  CUSTOM_0 As LCD_CHAR
  CUSTOM_1 As LCD_CHAR
  CUSTOM_2 As LCD_CHAR
  CUSTOM_3 As LCD_CHAR
  CUSTOM_4 As LCD_CHAR
  CUSTOM_5 As LCD_CHAR
  CUSTOM_6 As LCD_CHAR
  CUSTOM_7 As LCD_CHAR
End Type
Global LCD_CUSTOM As LCD_CUSTOM_i
Global Const LPT As Long = &H378
Global Const LCD_DATA_ADDRESS As Long = LPT + 0
Global Const LCD_CONTROL_ADDRESS As Long = LPT + 2
Global Const LCD_CHAR_BAR1 As Long = 0
Global Const LCD_CHAR_BAR2 As Long = 1
Global Const LCD_CHAR_BAR3 As Long = 2
Global Const LCD_CHAR_BAR4 As Long = 3
Global Const LCD_CHAR_BAR5 As Long = 4
Global Const LCD_CHAR_UP_ARROW As Long = 5
Global Const LCD_CHAR_DOWN_ARROW As Long = 6
Global Const BLOCK As String = "ÿ"
Public Sub LCD_LoadCustomChar()
'You may have up to 8 Custom Characters
'Addresses for Custom Characters are 0x00 through 0x07
'                                    &H0          &H7
'                                    Chr(0)       Chr(7)
'
LCD_CUSTOM.CUSTOM_0.ROW_0 = &H10 '  //  10000  //  M
LCD_CUSTOM.CUSTOM_0.ROW_1 = &H10 '  //  10000  //  M
LCD_CUSTOM.CUSTOM_0.ROW_2 = &H10 '  //  10000  //  M
LCD_CUSTOM.CUSTOM_0.ROW_3 = &H10 '  //  10000  //  M
LCD_CUSTOM.CUSTOM_0.ROW_4 = &H10 '  //  10000  //  M
LCD_CUSTOM.CUSTOM_0.ROW_5 = &H10 '  //  10000  //  M
LCD_CUSTOM.CUSTOM_0.ROW_6 = &H10 '  //  10000  //  M
LCD_CUSTOM.CUSTOM_0.ROW_7 = &H10 '  //  10000  //  M
'''
LCD_CUSTOM.CUSTOM_1.ROW_0 = &H18 '  //  11000  //  MM
LCD_CUSTOM.CUSTOM_1.ROW_1 = &H18 '  //  11000  //  MM
LCD_CUSTOM.CUSTOM_1.ROW_2 = &H18 '  //  11000  //  MM
LCD_CUSTOM.CUSTOM_1.ROW_3 = &H18 '  //  11000  //  MM
LCD_CUSTOM.CUSTOM_1.ROW_4 = &H18 '  //  11000  //  MM
LCD_CUSTOM.CUSTOM_1.ROW_5 = &H18 '  //  11000  //  MM
LCD_CUSTOM.CUSTOM_1.ROW_6 = &H18 '  //  11000  //  MM
LCD_CUSTOM.CUSTOM_1.ROW_7 = &H18 '  //  11000  //  MM
'''
LCD_CUSTOM.CUSTOM_2.ROW_0 = &H1C '  //  11100  //  MMM
LCD_CUSTOM.CUSTOM_2.ROW_1 = &H1C '  //  11100  //  MMM
LCD_CUSTOM.CUSTOM_2.ROW_2 = &H1C '  //  11100  //  MMM
LCD_CUSTOM.CUSTOM_2.ROW_3 = &H1C '  //  11100  //  MMM
LCD_CUSTOM.CUSTOM_2.ROW_4 = &H1C '  //  11100  //  MMM
LCD_CUSTOM.CUSTOM_2.ROW_5 = &H1C '  //  11100  //  MMM
LCD_CUSTOM.CUSTOM_2.ROW_6 = &H1C '  //  11100  //  MMM
LCD_CUSTOM.CUSTOM_2.ROW_7 = &H1C '  //  11100  //  MMM
'''
LCD_CUSTOM.CUSTOM_3.ROW_0 = &H1E '  //  11110  //  MMMM
LCD_CUSTOM.CUSTOM_3.ROW_1 = &H1E '  //  11110  //  MMMM
LCD_CUSTOM.CUSTOM_3.ROW_2 = &H1E '  //  11110  //  MMMM
LCD_CUSTOM.CUSTOM_3.ROW_3 = &H1E '  //  11110  //  MMMM
LCD_CUSTOM.CUSTOM_3.ROW_4 = &H1E '  //  11110  //  MMMM
LCD_CUSTOM.CUSTOM_3.ROW_5 = &H1E '  //  11110  //  MMMM
LCD_CUSTOM.CUSTOM_3.ROW_6 = &H1E '  //  11110  //  MMMM
LCD_CUSTOM.CUSTOM_3.ROW_7 = &H1E '  //  11110  //  MMMM
'''
LCD_CUSTOM.CUSTOM_4.ROW_0 = &H1F '  //  11111  //  MMMMM
LCD_CUSTOM.CUSTOM_4.ROW_1 = &H1F '  //  11111  //  MMMMM
LCD_CUSTOM.CUSTOM_4.ROW_2 = &H1F '  //  11111  //  MMMMM
LCD_CUSTOM.CUSTOM_4.ROW_3 = &H1F '  //  11111  //  MMMMM
LCD_CUSTOM.CUSTOM_4.ROW_4 = &H1F '  //  11111  //  MMMMM
LCD_CUSTOM.CUSTOM_4.ROW_5 = &H1F '  //  11111  //  MMMMM
LCD_CUSTOM.CUSTOM_4.ROW_6 = &H1F '  //  11111  //  MMMMM
LCD_CUSTOM.CUSTOM_4.ROW_7 = &H1F '  //  11111  //  MMMMM
'''Up Arrow
LCD_CUSTOM.CUSTOM_5.ROW_0 = &H1F '  //  11111  //  MMMMM
LCD_CUSTOM.CUSTOM_5.ROW_1 = &H1B '  //  11011  //  MM MM
LCD_CUSTOM.CUSTOM_5.ROW_2 = &H11 '  //  10001  //  M   M
LCD_CUSTOM.CUSTOM_5.ROW_3 = &HA '   //  01010  //   M M
LCD_CUSTOM.CUSTOM_5.ROW_4 = &H1B '  //  11011  //  MM MM
LCD_CUSTOM.CUSTOM_5.ROW_5 = &H1B '  //  11011  //  MM MM
LCD_CUSTOM.CUSTOM_5.ROW_6 = &H1B '  //  11011  //  MM MM
LCD_CUSTOM.CUSTOM_5.ROW_7 = &H1F '  //  11111  //  MMMMM
'''Down Arrow
LCD_CUSTOM.CUSTOM_6.ROW_0 = &H1F '  //  11111  //  MMMMM
LCD_CUSTOM.CUSTOM_6.ROW_1 = &H1B '  //  11011  //  MM MM
LCD_CUSTOM.CUSTOM_6.ROW_2 = &H1B '  //  11011  //  MM MM
LCD_CUSTOM.CUSTOM_6.ROW_3 = &H1B '  //  11011  //  MM MM
LCD_CUSTOM.CUSTOM_6.ROW_4 = &HA '   //  01010  //   M M
LCD_CUSTOM.CUSTOM_6.ROW_5 = &H11 '  //  10001  //  M   M
LCD_CUSTOM.CUSTOM_6.ROW_6 = &H1B '  //  11011  //  MM MM
LCD_CUSTOM.CUSTOM_6.ROW_7 = &H1F '  //  11111  //  MMMMM
End Sub
Public Sub LCD_Init()
'LCD_Init: Initialize the LCD.
LCD_InitDriver
LCD_Clear
LCD_CursorOff
LCD_LoadCustomChar
LCD_DefineChar &H0, LCD_CUSTOM.CUSTOM_0
LCD_DefineChar &H1, LCD_CUSTOM.CUSTOM_1
LCD_DefineChar &H2, LCD_CUSTOM.CUSTOM_2
LCD_DefineChar &H3, LCD_CUSTOM.CUSTOM_3
LCD_DefineChar &H4, LCD_CUSTOM.CUSTOM_4
LCD_DefineChar &H5, LCD_CUSTOM.CUSTOM_5
LCD_DefineChar &H6, LCD_CUSTOM.CUSTOM_6
LCD_Home
End Sub
Public Sub LCD_Clear()
'LCD_Clear: Clear the LCD screen (also homes cursor).
LCD_WriteControl &H1
End Sub
Public Sub LCD_Home()
'LCD_Home: Position the LCD cursor at row 1, col 1.
LCD_Cursor 1, 1
End Sub
Public Sub LCD_DisplayCharacter(aChar As String)
'LCD_DisplayCharacter: Display a single character, at the current cursor location.
LCD_WriteData Asc(aChar)
End Sub
Public Sub LCD_DisplayString(Row As Integer, Column As Integer, aString As String)
'LCD_DisplayString: Display a string at the specified row and column.
Dim I As Integer
LCD_Cursor Row, Column
For I = 0 To Len(aString) - 1
  LCD_DisplayCharacter Mid(aString, I + 1, 1)
Next I
End Sub
Public Sub LCD_DisplayStringCentered(Row As Integer, aString As String)
'LCD_DisplayStringCentered: Display a string centered on the specified row.
Dim N As Integer
Dim I As Integer
N = Len(aString)
If N <= 20 Then
  LCD_Cursor Row, 1
  For I = 0 To 19
     LCD_DisplayCharacter " "
  Next I
  LCD_DisplayString Row, ((20 - N) / 2) + 1, aString
Else
  LCD_DisplayString Row, 1, aString
End If
End Sub
Public Sub LCD_Cursor(Row As Integer, Column As Integer)
'LCD_Cursor: Position the LCD cursor at "row", "column".
Select Case Row
  Case 1
    LCD_WriteControl &H80 + Column - 1
  Case 2
    LCD_WriteControl &HC0 + Column - 1
  Case 3
    LCD_WriteControl &H94 + Column - 1
  Case 4
    LCD_WriteControl &HD4 + Column - 1
End Select
End Sub
Public Sub LCD_DisplayScreen(PTR As String)
'LCD_DisplayScreen: Display an entire screen (80 characters).
'inputs: PTR = A string containing the entire screen
'example:
'Test = "01234567890123456789"_
'      &" This is a test of  "_
'      &"LCD_DisplayScreen()."_
'      &"   How's it look?   "
'LCD_DisplayScreen Test
LCD_DisplayRow 1, Left(PTR, 20)
LCD_DisplayRow 2, Mid(PTR, 21, 20)
LCD_DisplayRow 3, Mid(PTR, 41, 20)
LCD_DisplayRow 4, Right(PTR, 20)
End Sub
Public Sub LCD_WipeOnLR(PTR As String)
'LCD_WipeOnLR: Display an entire screen (80 characters) by "wiping" it on (left to right).
'inputs: PTR = A string containing the entire screen.
Dim I As Integer
For I = 1 To 20
  LCD_Cursor 1, I
  LCD_DisplayCharacter Mid(PTR, I, I)
  LCD_Cursor 2, I
  LCD_DisplayCharacter Mid(PTR, I + 20, I)
  LCD_Cursor 3, I
  LCD_DisplayCharacter Mid(PTR, I + 40, I)
  LCD_Cursor 4, I
  LCD_DisplayCharacter Mid(PTR, I + 60, I)
Next I
End Sub
Public Sub LCD_WipeOnRL(PTR As String)
'LCD_WipeOnLR: Display an entire screen (80 characters) by "wiping" it on (right to left).
'inputs: PTR = A string containing the entire screen.
Dim I As Integer
For I = 20 To 1 Step -1
  LCD_Cursor 1, I
  LCD_DisplayCharacter Mid(PTR, I, I)
  LCD_Cursor 2, I
  LCD_DisplayCharacter Mid(PTR, 20 + I, I)
  LCD_Cursor 3, I
  LCD_DisplayCharacter Mid(PTR, 40 + I, I)
  LCD_Cursor 4, I
  LCD_DisplayCharacter Mid(PTR, 60 + I, I)
Next I
End Sub
Public Sub LCD_WipeOffLR()
'LCD_WipeOffLR: "Wipe" screen left-to-right.
Dim I As Integer
For I = 1 To 20
  LCD_Cursor 1, I
  LCD_DisplayCharacter BLOCK
  LCD_Cursor 2, I
  LCD_DisplayCharacter BLOCK
  LCD_Cursor 3, I
  LCD_DisplayCharacter BLOCK
  LCD_Cursor 4, I
  LCD_DisplayCharacter BLOCK
Next I
End Sub
Public Sub LCD_WipeOffRL()
'LCD_WipeOffRL: "Wipe" screen right-to-left.
Dim I As Integer
For I = 20 To 1 Step -1
  LCD_Cursor 1, I
  LCD_DisplayCharacter BLOCK
  LCD_Cursor 2, I
  LCD_DisplayCharacter BLOCK
  LCD_Cursor 3, I
  LCD_DisplayCharacter BLOCK
  LCD_Cursor 4, I
  LCD_DisplayCharacter BLOCK
Next I
End Sub
Public Sub LCD_DisplayRow(Row As Integer, aString As String)
'LCD_DisplayRow: Display a string at the specified row.
Dim I As Integer
LCD_Cursor Row, 1
For I = 1 To 20
  LCD_DisplayCharacter Mid(aString, I, 1)
Next I
End Sub
Public Sub LCD_CursorLeft()
'LCD_CursorLeft: Move the cursor left by one character.
LCD_WriteControl &H10
End Sub
Public Sub LCD_CursorRight()
'LCD_CursorRight: Move the cursor right by one character.
LCD_WriteControl &H14
End Sub
Public Sub LCD_CursorOn()
'LCD_CursorOn: Turn the cursor on.
LCD_WriteControl &HD
End Sub
Public Sub LCD_CursorOff()
'LCD_CursorOff: Turn the cursor off.
LCD_WriteControl &HC
End Sub
Public Sub LCD_DisplayOff()
'LCD_DisplayOff: Turn Off LCD.
LCD_WriteControl &H8
End Sub
Public Sub LCD_DisplayOn()
'LCD_DisplayOn: Turn On LCD.
LCD_WriteControl &HC
End Sub
Private Sub LCD_InitDriver()
'LCD_InitDriver: Initialize the LCD driver.
LCD_WriteControl &H38
LCD_WriteControl &H38
LCD_WriteControl &H38
LCD_WriteControl &H6
LCD_WriteControl &HC
End Sub
Private Sub LCD_WriteControl(Data As Long)
'LCD_WriteControl: Write a control instruction to the LCD
vbOut LCD_CONTROL_ADDRESS, &H3 ' RS=0, R/W=0, E=0
vbOut LCD_DATA_ADDRESS, Data
vbOut LCD_CONTROL_ADDRESS, &H2 ' RS=0, R/W=0, E=1
vbOut LCD_CONTROL_ADDRESS, &H3 ' RS=0, R/W=0, E=0
vbOut LCD_CONTROL_ADDRESS, &H1 ' RS=0, R/W=1, E=0
Sleep 10
End Sub
Private Sub LCD_WriteData(Data As Long)
'LCD_WriteData: Write one byte of data to the LCD
vbOut LCD_CONTROL_ADDRESS, &H7 ' RS=1, R/W=0, E=0
vbOut LCD_DATA_ADDRESS, Data
vbOut LCD_CONTROL_ADDRESS, &H6 ' RS=1, R/W=0, E=1
vbOut LCD_CONTROL_ADDRESS, &H7 ' RS=1, R/W=0, E=0
vbOut LCD_CONTROL_ADDRESS, &H5 ' RS=1, R/W=1, E=0
Sleep 1
End Sub
Public Sub LCD_DefineChar(Address As Long, Pattern As LCD_CHAR)
'LCD_DefineCharacter: Define dot pattern for user-defined character.
'inputs: Address = address of character (0x00-0x07)
'        Pattern = pointer to 8-byte array containing the dot pattern
Dim I As Integer
LCD_WriteControl &H40 + vbShiftLeft(Address, 3)
LCD_WriteData Pattern.ROW_0
LCD_WriteData Pattern.ROW_1
LCD_WriteData Pattern.ROW_2
LCD_WriteData Pattern.ROW_3
LCD_WriteData Pattern.ROW_4
LCD_WriteData Pattern.ROW_5
LCD_WriteData Pattern.ROW_6
LCD_WriteData Pattern.ROW_7
End Sub
Public Function vbShiftLeft(ByVal Value As Long, Count As Integer) As Long
'This function is equivalent to the 'C' language construct '<<'
Dim I As Integer
vbShiftLeft = Value
For I = 1 To Count
  vbShiftLeft = vbShiftLeft * 2
Next
End Function
