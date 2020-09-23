VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   Caption         =   "Falling Code Display Window"
   ClientHeight    =   2880
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2880
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   1800
      Top             =   2040
   End
   Begin VB.Timer Timer1 
      Left            =   1800
      Top             =   1440
   End
   Begin VB.PictureBox Code 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   20000
      Index           =   0
      Left            =   1080
      ScaleHeight     =   19935
      ScaleWidth      =   225
      TabIndex        =   0
      Top             =   120
      Width           =   285
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
   End
   Begin VB.Menu mnuplypaus 
      Caption         =   "&Pause"
   End
   Begin VB.Menu mnuabout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public GRed As Long, GGreen As Long, GBlue As Long
Private Sub Drawtext(Index As Long)
Dim Red1 As Integer, Red2 As Integer, Green1 As Integer, Green2 As Integer, Blue1 As Integer, Blue2 As Integer, RedInt As Integer, GreenInt As Integer, BlueInt As Integer, MaxLen As Integer, Counter As Integer
'Set the color to fade to. In this case the form is black so we want to fade to black (0,0,0)
Red1 = 0
Green1 = 0
Blue1 = 0
'Set the color we want to fade from. In this case, the RGB values of this color are determined by scrollbars on the edit form
Red2 = GRed
Green2 = GGreen
Blue2 = GBlue
'Clear out the picturebox so that everything we add to it will be visible.
Code(Index).Cls
'Use the Rnd variable to determine a random legnth we want the line of code to be (between 3 and 250)
MaxLen = (Rnd * 250) + 3
'Set the height of the picturebox so that it comes close to matching the approximate height of the text within it
Code(Index).Height = 15 * Screen.TwipsPerPixelY * MaxLen
'Get the absolute value of the difference (to avaoid errors with negative numbers) between the colors to fade from and to fade to, and then divide them by the number of characters so that we can get a color change interval (subtract 1 so that there is no color change on the first letter)
RedInt = Abs(Red2 - Red1) \ (MaxLen - 1)
GreenInt = Abs(Green2 - Green1) \ (MaxLen - 1)
BlueInt = Abs(Blue2 - Blue1) \ (MaxLen - 1)
'Draw the text into the picturebox
For Counter = 0 To MaxLen
 'Set the forcolor of the picturebox to reflect the line of the faded color so that when the text is printed in the picturebox, it will be the correct color.
 Code(Index).ForeColor = RGB(Red1 + (Counter * RedInt), Green1 + (Counter * GreenInt), Blue1 + (Counter * BlueInt))
 'Print a random character in the ASCII range [161-255] into the picturebox
 Code(Index).Print Chr((Rnd * 94) + 161)
Next Counter
'Set the tag of the picture box of a character with an ASCII value in the range [1-50] to designate the speed in wich the line of "code" will fall of the screen.
Code(Index).Tag = Chr(Int(Rnd * 30) + 1)
'Move the picture box right above the top of the form so that we can watch it fall down the form
Code(Index).Top = -Code(Index).Height
End Sub
Private Sub MakeNewCode()
Dim Counter As Long, Token As Long, MaxLines As Long
'Make it so that the program does not halt for a second while redrawing the code on the screen
DoEvents
'Check to make sure that child pictureboxes exist, so that we do not get an error from trying to unload controls that have not been loaded
If Code.UBound > 0 Then
 For Counter = Code.UBound To 1 Step -1
  'Unload the picturebox to avoid errors when the program tries to reload the picturebox again
  Unload Code(Counter)
 Next Counter
End If
'Determine how many "lines" of code there will be by dividing the width of the form by the width of the picturebox
MaxLines = Me.Width \ Code(0).Width + 1
For Token = 1 To MaxLines
 'Create a new picturebox
 Load Code(Token)
 'Set the left of the picturebox relative to its indx, so that all the lines of code will appear side by side, rather than just one big line that keeps overlapping itself
 Code(Token).Left = (Token - 1) * Code(Token).Width
 'Show the picturebox, so that the person watching the program will be able to see the scrolling text
 Code(Token).Visible = True
 'Draw the contents of the picture box and shift it to a position so that it will "fall" down the screen.
 Drawtext Token
Next Token
End Sub

Private Sub Form_Load()
'Code 0 is only supposed to be used as a template. so we dont want the user to see it during runtime.
Code(0).Visible = False
'Set the picturebox's border style to none, as we dont want the code to be contained in falling boxes.
Code(0).BorderStyle = 0
'Set the Code 0 tag to a useless value, just so that we do not encounter any errors during runtime when the timed checks for a value in it.
Code(0).Tag = "a"
'Set the picturebox's autoredraw value to true, so that the date in the pictureboxes will not be erased until the program tells it to do so
Code(0).AutoRedraw = True
'Set the backcolor of the picturebox to that of the form, so that the text isn't surrounded by gray boxes
Code(0).BackColor = Me.BackColor
'Set the timeer intrval for 10, so that the refresh will be fast and smooth
Timer1.Interval = 1
'Set the value of the global variables so that we have a color for the code to start scrolling in
GRed = 0
GGreen = 240
GBlue = 0
'Create the textboxes (lines of "code") that will fill up the form
MakeNewCode
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Close the program when the main form is unloaded. This keeps it from running hidden in the background after the form is closed.
End
End Sub

Private Sub mnuabout_Click()
MsgBox "Falling code - Matrix style text effect demonstration by Michael Dzicek", , "About the Falling Code demonstration"
End Sub

Private Sub mnuEdit_Click()
frmEdit.Show
End Sub

Private Sub mnuplypaus_Click()
If Timer1.Enabled = True Then
 'Set the button's caption to Play
 mnuplypaus.Caption = "&Play"
 'Halt the timer that causes the code to move
 Timer1.Enabled = False
Else
 'Set the button's caption to Pause
 mnuplypaus.Caption = "&Pause"
 'Start the timer back up so that the code moves again
 Timer1.Enabled = True
End If
End Sub

Private Sub Timer1_Timer()
Dim Token As Long
'Run through every picturebox (via its index number) to update it
For Token = Code.lBound To Code.UBound
 'Decode the value of its tag, and translate it into its speed, which the picturebox is then shifted down
 Code(Token).Top = Code(Token).Top + (Asc(Left(Code(Token).Tag, 1)) * Screen.TwipsPerPixelY)
 'Check to make sure that the picturebox has not "fallen" offscreen. If so, redraw its contents, and shift it back up to the top of the form.
 If Code(Token).Top > Me.Height Then Drawtext Token
Next Token
End Sub

Private Sub Timer2_Timer()
'Draw the code to match the form size in the case that the form should be adjusted. Would have put this in the form resize event, but visualbasic has a problem with putting the unload statement into the resize event.
'Also its in a second timer so that it does not slow down the progress of the text scrolling as much as it would have otherwise.
If (Code(Code.UBound).Left + Code(Code.UBound).Width * 2) < Me.Width Then MakeNewCode
End Sub
