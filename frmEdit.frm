VERSION 5.00
Begin VB.Form frmEdit 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Code Properties"
   ClientHeight    =   1185
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   3510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar Blue 
      Height          =   975
      Left            =   2040
      Max             =   255
      TabIndex        =   3
      Top             =   0
      Width           =   255
   End
   Begin VB.VScrollBar Green 
      Height          =   975
      Left            =   1200
      Max             =   255
      TabIndex        =   2
      Top             =   0
      Width           =   255
   End
   Begin VB.VScrollBar Red 
      Height          =   975
      Left            =   360
      Max             =   255
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox ColorProp 
      Height          =   855
      Left            =   2640
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Code Color:"
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      Top             =   0
      Width           =   855
   End
   Begin VB.Label BlueProp 
      Alignment       =   2  'Center
      Caption         =   "Label3"
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   960
      Width           =   735
   End
   Begin VB.Label GreenProp 
      Alignment       =   2  'Center
      Caption         =   "Label2"
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   960
      Width           =   855
   End
   Begin VB.Label RedProp 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   960
      Width           =   855
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ClearColVals()
'change all scroll bar values to 1 so that the captions update
Red.Value = 1
Green.Value = 1
Blue.Value = 1
End Sub

Private Sub Blue_Change()
'Set the color of the picturebox to represent the colors dictated by the scrollbar values
ColorProp.BackColor = RGB(Red.Value, Green.Value, Blue.Value)
'Make the blue caption show the blue value
BlueProp.Caption = "Blue: " & Blue.Value
End Sub

Private Sub Blue_Scroll()
'Set the color of the picturebox to represent the colors dictated by the scrollbar values
ColorProp.BackColor = RGB(Red.Value, Green.Value, Blue.Value)
'Make the blue caption show the blue value
BlueProp.Caption = "Blue: " & Blue.Value
End Sub

Private Sub Form_Load()
'update the labels so that they will read the values of the global variables
ClearColVals
'set the value of each of the color scroll bars to match the colors of the corresponding global variable
Red.Value = frmMain.GRed
Green.Value = frmMain.GGreen
Blue.Value = frmMain.GBlue
End Sub

Private Sub Form_Unload(Cancel As Integer)
'set the "falling code" colors to the colors that were defined by this edit form
frmMain.GRed = Red.Value
frmMain.GGreen = Green.Value
frmMain.GBlue = Blue.Value
End Sub

Private Sub Green_Change()
'Set the color of the picturebox to represent the colors dictated by the scrollbar values
ColorProp.BackColor = RGB(Red.Value, Green.Value, Blue.Value)
'Make the green caption show the green value
GreenProp.Caption = "Green: " & Green.Value
End Sub

Private Sub Green_Scroll()
'Set the color of the picturebox to represent the colors dictated by the scrollbar values
ColorProp.BackColor = RGB(Red.Value, Green.Value, Blue.Value)
'Make the green caption show the green value
GreenProp.Caption = "Green: " & Green.Value
End Sub

Private Sub Red_Change()
'Set the color of the picturebox to represent the colors dictated by the scrollbar values
ColorProp.BackColor = RGB(Red.Value, Green.Value, Blue.Value)
'Make the red caption show the red value
RedProp.Caption = "Red: " & Red.Value
End Sub

Private Sub Red_Scroll()
'Set the color of the picturebox to represent the colors dictated by the scrollbar values
ColorProp.BackColor = RGB(Red.Value, Green.Value, Blue.Value)
'Make the red caption show the red value
RedProp.Caption = "Red: " & Red.Value
End Sub
