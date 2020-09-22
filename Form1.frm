VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "EM_LINESCROLL Demo"
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   ScaleHeight     =   447
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   408
   StartUpPosition =   2  'CenterScreen
   Begin VB.VScrollBar VScroll2 
      Height          =   5655
      Left            =   5640
      TabIndex        =   3
      Top             =   360
      Width           =   255
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   5655
      Left            =   2640
      TabIndex        =   2
      Top             =   360
      Width           =   255
   End
   Begin RichTextLib.RichTextBox RTB2 
      Height          =   5655
      Left            =   3120
      TabIndex        =   1
      Top             =   360
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   9975
      _Version        =   393217
      TextRTF         =   $"Form1.frx":0000
   End
   Begin RichTextLib.RichTextBox RTB1 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   9975
      _Version        =   393217
      TextRTF         =   $"Form1.frx":00E1
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   6360
      Width           =   5775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select A Test To Do From The Menu To Start"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   6120
      Width           =   5895
   End
   Begin VB.Menu mnuDemoMain 
      Caption         =   "Demos"
      Begin VB.Menu mnuDemo2 
         Caption         =   "Demo 2 - Simultaneous Scrolling Precision Line Positioning"
      End
   End
   Begin VB.Menu mnuTestsMain 
      Caption         =   "Tests To Do"
      Begin VB.Menu mnuT1ArrowClick 
         Caption         =   "Test #1: Arrow Click"
      End
      Begin VB.Menu mnuT2SlowDrag 
         Caption         =   "Test #2: Slow  Small Drag"
      End
      Begin VB.Menu mnuT3LargeDrag 
         Caption         =   "Test #3: Large Drag"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'I hate this but I do it for you!
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++ EM_LINESCROLL Demo by CptnVic                                                      ++
'++ Submitted to Planet Source Code...                                                 ++
'++ If you found it elsewhere... it was hoarked from Planet Source Code!               ++
'++ So shoot the bastar... er... frown at the hoser who done it!                       ++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++ I HOPE that you don't come away with the idea that I am ragging on KRYO_11's       ++
'++ code.  I am not!  His (her... I don't know) willingness to share code with us      ++
'++ is the entire reason that this one was written.  I was completely satisfied        ++
'++ with the results I was getting using KRYO_11's code until I put it to use on       ++
'++ a rich textbox with lots of lines in it.  Thanks to KRYO_11 for the idea!          ++
'++ You can see his code at:                                                           ++
'++ http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=54152&lngWId=1 ++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++ It took me about half a day to track down this constant and get a bug free         ++
'++ implementation of this api to work... mostly because I spent too much              ++
'++ time over engineering the very simple.  I'm stupid like that at times!             ++
'++ When I finally got it boiled down to it's simplest form... it is easy to use.      ++
'++ IF you do find the code useful, a little credit would be nice... maybe a vote      ++
'++ or a NICE comment... OTHERWISE... You have to make me Emperor of the universe      ++
'++ and COMPLETE AND PAY my income taxes to use this code.  ;-)                        ++
'++ I hope you will find this helpful!                                                 ++
'++ Your most ardent admirer and partner in crime,                                     ++
'++ CptnVic                                                                            ++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'declare the api functions
'don't need the following declaration in my version
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'needed in my version:
Private Declare Function SendMessageBynum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)

'declare some constants for sendmessage
Private Const EM_LINESCROLL = &HB6 'needed in my version
Private Const EM_SCROLL As Long = &HB5 'needed in both versions
Private Const EM_GETLINECOUNT As Long = &HBA 'needed in both versions


Dim PSP As Integer 'used for RTB1
Dim LastScrollVal As Integer 'used to store last scrollbar value for RTB2

Private Sub Form_Load()
    Dim Junk As String, X As Integer
    RTB1.Text = "": RTB2.Text = 0 'clear both text boxes
    'put some text into both RTB's... so we can get a line count
    Junk = ""
    For X = 1 To 1000
        Junk = Junk & "This is line: " & Str(X) & vbCrLf
    Next
    RTB1.Text = Junk
    RTB2.Text = Junk
    'set the scroll bar values... make it a fair test... they are the same
    SetUpRTB1
    SetUpRTB2
    'print a brief instruction
    Me.CurrentX = RTB1.Left
    Me.CurrentY = 10
    Me.Print "RTB1 (KRYO_11's method)";
    Me.CurrentX = RTB2.Left
    Me.Print "RTB2 (CptnVic's method)"
    Label2.Caption = ""
End Sub
Private Sub SetUpRTB1()
    'this is piped in from the original form_load... only changed to reflect the test rtb's and pixel scale mode of the form
    'RichTextBox1.Text = Text1.Text 'it's already populated with text
    Dim lCount As Long
    lCount = SendMessage(RTB1.hwnd, EM_GETLINECOUNT, 0, ByVal 0&) 'retrieve # total of lines
    VScroll1.Min = 0
    VScroll1.Max = lCount - ((RTB1.Height) / Me.TextHeight("A")) 'adjust according to height of a line... I changed this just a bit because I'm using pixel scale mode to print the labels
    VScroll1.Value = 0
    VScroll1.SmallChange = 1
    VScroll1.LargeChange = 10
    PSP = 0
End Sub
Private Sub SetUpRTB2()
'set up 2nd rtb same as the first
    Dim lCount As Long
    lCount = SendMessageBynum(RTB2.hwnd, EM_GETLINECOUNT, 0, ByVal 0&)
    VScroll2.Min = 0
    VScroll2.Max = lCount - ((RTB2.Height) / Me.TextHeight("A")) 'adjust according to height of a line
    'BTW: EM_LINESCROLL will not scroll past the last line in the text box... no matter what values are passed to it
    'so if you wrote this line as:
    'VScroll2.Max = lCount ' or VScroll2.Max = lCount + 25000...
    'VScroll2.Max = lCount + 25000 'try it if you don't believe me
    'and did not adjust... you'd see empty lines below the last line that don't "really" exist...
    'and are not included in the line count returned by EM_GETLINECOUNT
    'but that's insane... and pointless... it will just screw up your scroll bar values!
    VScroll2.Value = 0
    VScroll2.SmallChange = 1
    VScroll2.LargeChange = 10
    LastScrollVal = 0
End Sub

Private Sub mnuDemo2_Click()
    Form2.Show
End Sub

Private Sub mnuT1ArrowClick_Click()
    Label1.ForeColor = &HFF0000
    Label1.Caption = "Click The Down Arrows On BOTH Scroll Bars Several Times"
    Label2.Caption = "You shouldn't notice much difference between the two text boxes"
End Sub

Private Sub mnuT2SlowDrag_Click()
    Label1.ForeColor = &H0
    Label1.Caption = "Drag The Scroll Bar Handles On BOTH Scroll Bars SLOWLY To Line 1"
    Label2.Caption = "You shouldn't notice much difference between the two text boxes"
End Sub

Private Sub mnuT3LargeDrag_Click()
    Label1.ForeColor = &HFF0000
    Label1.Caption = "Drag The Scroll Bar Handles On BOTH Scroll Bars QUICKLY To The Maximum"
    Label2.Caption = "Did You notice any difference between the two text boxes?"
End Sub

Private Sub VScroll1_Change()
    'pass this thru to _Scroll event to save code
    VScroll1_Scroll
End Sub

Private Sub VScroll1_Scroll()
'piped in from KRYO_11's code
    Dim l As Long
    Dim i
    With RTB1
            
        If VScroll1.Value > PSP Then
            For i = PSP + 1 To VScroll1.Value
                    l = SendMessage(.hwnd, EM_SCROLL, 1, 0)
            Next i
        ElseIf VScroll1.Value < PSP Then
            For i = VScroll1.Value + 1 To PSP
                l = SendMessage(.hwnd, EM_SCROLL, 0, 1)
            Next i
        End If

        PSP = VScroll1.Value
        
    End With
End Sub

Private Sub VScroll2_Change()
    'pass this thru to _Scroll event to save code
    VScroll2_Scroll
End Sub

Private Sub VScroll2_Scroll()
'for some reason... people seem to be interested in simultaneously scrolling multiple
'text boxes... this code will do exactly that... see the form2 demo
'this code is so fast it causes some flicker if the change is large...
    Dim Vert As Integer
        If VScroll2.Value > LastScrollVal Then 'scroll down
            Vert = VScroll2.Value - LastScrollVal 'calc the change needed
            SendMessageBynum RTB2.hwnd, EM_LINESCROLL, 0, Vert 'do the scrolling
        ElseIf VScroll2.Value < LastScrollVal Then 'scroll up
            Vert = LastScrollVal - VScroll2.Value 'calc the change needed
            SendMessageBynum RTB2.hwnd, EM_LINESCROLL, 0, -Vert 'do the scrolling
        End If
    LastScrollVal = VScroll2.Value 'store the current value for next change
End Sub
