VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   Caption         =   "EM_LINESCROLL Demo #2   Simultaneous Scroll and Line Number Synchronization"
   ClientHeight    =   4395
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11505
   LinkTopic       =   "Form2"
   ScaleHeight     =   293
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   767
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox SimRTB 
      Height          =   2775
      Index           =   2
      Left            =   9720
      TabIndex        =   10
      Top             =   1440
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   4895
      _Version        =   393217
      Enabled         =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Form2.frx":0000
   End
   Begin RichTextLib.RichTextBox SimRTB 
      Height          =   2775
      Index           =   3
      Left            =   9240
      TabIndex        =   9
      Top             =   1440
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   4895
      _Version        =   393217
      BackColor       =   14737632
      Enabled         =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Form2.frx":00E1
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1080
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   720
      Width           =   3855
   End
   Begin VB.CommandButton cmdGotoLineNum 
      Caption         =   "GoTo That Line"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5160
      TabIndex        =   7
      Top             =   720
      Width           =   1455
   End
   Begin VB.CheckBox chkboxNoSync 
      Caption         =   "Disable Line Synchronization"
      Height          =   255
      Left            =   5160
      TabIndex        =   6
      Top             =   120
      Width           =   4335
   End
   Begin VB.CheckBox chkboxNoLBScroll 
      Caption         =   "Disable List Box Scrolling"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   120
      Width           =   3975
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   6960
      TabIndex        =   4
      Top             =   1440
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   2775
      Left            =   4680
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "Form2.frx":01C2
      Top             =   1440
      Width           =   2055
   End
   Begin RichTextLib.RichTextBox SimRTB 
      Height          =   2775
      Index           =   1
      Left            =   2520
      TabIndex        =   2
      Top             =   1440
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   4895
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form2.frx":01C8
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2775
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   255
   End
   Begin RichTextLib.RichTextBox SimRTB 
      Height          =   2775
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   1440
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   4895
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form2.frx":02A9
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Code Box/Line # Example"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9240
      TabIndex        =   11
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Menu mnuTestsMain 
      Caption         =   "Tests To Do"
      Begin VB.Menu mnuT1SimulScroll 
         Caption         =   "Test #1: Simultaneous Scrolling In Sync"
      End
      Begin VB.Menu mnuT2NoListBox 
         Caption         =   "Test #2: Turn Off List Box Scrolling"
      End
      Begin VB.Menu mnuT3ResyncList 
         Caption         =   "Test #3: Re-Synchronize List Box"
      End
      Begin VB.Menu mnuT4NoSync 
         Caption         =   "Test #4: Un-Synchronized Scrolling"
      End
      Begin VB.Menu mnuT5ReSync 
         Caption         =   "Test #5: Re-Synchronize Scrolling"
      End
      Begin VB.Menu mnuT6GotoLine 
         Caption         =   "Test #6: Precision Goto Line # Test"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'I hate this but I do it for you!
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++ EM_LINESCROLL Demo by CptnVic                                                  ++
'++ Submitted to Planet Source Code...                                             ++
'++ If you found it elsewhere... it was hoarked from Planet Source Code!           ++
'++ So shoot the bastar... er... frown at the hoser who done it!                   ++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++ It took me about half a day to track down this constant and get a bug free     ++
'++ implementation of this api to work... mostly because I spent too much          ++
'++ time over engineering the very simple.  I'm stupid like that at times!         ++
'++ When I finally got it boiled down to it's simplest form... it is easy to use.  ++
'++ IF you do find the code useful, a little credit would be nice... maybe a vote  ++
'++ or a NICE comment... OTHERWISE... You have to make me Emperor of the Universe  ++
'++ and COMPLETE AND PAY my income taxes to use this code.  ;-)                    ++
'++ I hope you will find this helpful!                                             ++
'++ Your most ardent admirer and partner in crime,                                 ++
'++ CptnVic                                                                        ++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'needed in my version: ... this could/should be in a module
Private Declare Function SendMessageBynum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)

'declare some constants for sendmessage
Private Const EM_LINESCROLL = &HB6 'needed in my version
Private Const EM_SCROLL As Long = &HB5 'needed in both versions
Private Const EM_GETLINECOUNT As Long = &HBA 'needed in both versions
Private Const EM_GETFIRSTVISIBLELINE = &HCE 'used to re-sync lines, returns topmost visible line #
'list box top constants
Private Const LB_GETTOPINDEX = &H18E
Private Const LB_SETTOPINDEX = &H197

Dim LastScrollVal As Integer 'used to store last scrollbar value for all objects
Dim Msg As String, Title As String 'msg box vars

Private Sub cmdGotoLineNum_Click()
    'jump to the selected line
    Dim FirstVis As Long, VSBVal As Integer, Dif As Integer, X As Integer
    VSBVal = Val(Combo1.Text) - 1 'first line = 0 so adjust
    If VSBVal > 100 Then
        Msg = "This is crazy!  There are only 100 Lines!" & vbCrLf & "But do it anyway... it won't crash!"
        MsgBox Msg
    End If
    For X = 0 To 3 'do the richtext boxes
        FirstVis = SendMessageBynum(SimRTB(X).hwnd, EM_GETFIRSTVISIBLELINE, 0, 0) 'get the topmost visible line #
        If FirstVis <> VSBVal Then
            Dif = VSBVal - FirstVis 'calculate the change needed
            SendMessageBynum SimRTB(X).hwnd, EM_LINESCROLL, 0, Dif 'scroll there
        End If
    Next
    'do it again for the text box
    FirstVis = SendMessageBynum(Text1.hwnd, EM_GETFIRSTVISIBLELINE, 0, 0)
    If FirstVis <> VSBVal Then
        Dif = VSBVal - FirstVis
        SendMessageBynum Text1.hwnd, EM_LINESCROLL, 0, Dif
    End If
    'sync-up the listbox
    If chkboxNoLBScroll.Value = 0 Then 'listbox scroll is on
        If VSBVal > List1.ListCount Then VSBVal = VScroll1.Max + 15 'prevent "Nothing" from happening
        'LB_SETTOPINDEX won't freak out at a bad value (like this situation)... it just does nothing.
        SendMessageBynum List1.hwnd, LB_SETTOPINDEX, VSBVal, 0 'do the scrolling
    End If
End Sub

Private Sub Combo1_Change()
    'this is an effort to keep you from typing in crazy stuff... it's pointless
    If Val(Combo1.Text) <= 0 Then
        cmdGotoLineNum.Enabled = False
    Else
        cmdGotoLineNum.Enabled = True
    End If
End Sub

Private Sub Combo1_Click()
    If Val(Combo1.Text) > 0 Then 'check to see if a number is selected before jumping
        cmdGotoLineNum.Enabled = True
    Else
        cmdGotoLineNum.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    'print some labels
    Me.CurrentY = SimRTB(0).Top - Me.TextHeight("W") - 5
    Me.CurrentX = SimRTB(0).Left: Me.Print "Rich Text Box: SimRTB(0)";
    Me.CurrentX = SimRTB(1).Left: Me.Print "Rich Text Box: SimRTB(1)";
    Me.CurrentX = Text1.Left: Me.Print "Text Box: Text1";
    Me.CurrentX = List1.Left: Me.Print "List Box: List1"
    'put stuff everywhere
    Dim X As Integer, Junk As String, ListJunk As String, CodeLine As String
    List1.Clear 'clear the listbox
    For X = 0 To 3
        SimRTB(X).Text = "" 'clear the text boxes
    Next
    'populate the things
    Combo1.Clear 'clear the combobox
    For X = 1 To 100
        ListJunk = "This is line: " & Str(X)
        Junk = Junk & "This is line: " & Str(X) & vbCrLf
        CodeLine = CodeLine & "Code line # " & Str(X) & vbCrLf
        List1.AddItem ListJunk
        Combo1.AddItem Str(X)
        SimRTB(3).Text = SimRTB(3).Text & Str(X) & vbCrLf
    Next
    SimRTB(0).Text = Junk
    SimRTB(1).Text = Junk
    SimRTB(2).Text = CodeLine
    'some bonus bogus choice for Combo1
    Combo1.AddItem Str(200)
    'add buffer space at the bottom of text1 & list1... so they will scroll far enough
    Text1.Text = Junk
    Junk = "" 'clear
    For X = 1 To 40
        Junk = Junk & vbCrLf
        List1.AddItem "" 'add empty items to the list
    Next
    Text1.Text = Text1.Text & Junk 'add empty items to the text
    SetUpScrollBar 'initialize the scrollbar
    Combo1.Text = "Pick A Line #"
End Sub

Private Sub SetUpScrollBar()
'set up the single scroll bar
    Dim lCount As Long
    lCount = SendMessageBynum(SimRTB(0).hwnd, EM_GETLINECOUNT, 0, ByVal 0&) 'get # of line to potentially scroll
    VScroll1.Min = 0
    VScroll1.Max = lCount - ((SimRTB(0).Height) / Me.TextHeight("A")) 'adjust according to height of a line
    VScroll1.Value = 0
    VScroll1.SmallChange = 1
    VScroll1.LargeChange = 10
    LastScrollVal = 0
End Sub
Private Sub ScrollStuff(Thing As Object)
'for some reason... people seem to be interested in simultaneously scrolling multiple
'text boxes... this code will do exactly that... and more
'this code is so fast it causes some flicker if the change is large...
'This sub would be very useful in adding a line number text box to a code editor for scrolling in sync.
    Dim Vert As Integer
        If VScroll1.Value > LastScrollVal Then 'scroll down
            Vert = VScroll1.Value - LastScrollVal 'calc the change needed
            SendMessageBynum Thing.hwnd, EM_LINESCROLL, 0, Vert 'do the scrolling
            If chkboxNoLBScroll.Value = 0 Then 'listbox scroll is on
                SendMessageBynum Thing.hwnd, LB_SETTOPINDEX, VScroll1.Value, 0 'do the scrolling
            End If
        ElseIf VScroll1.Value < LastScrollVal Then 'scroll up
            Vert = LastScrollVal - VScroll1.Value 'calc the change needed
            SendMessageBynum Thing.hwnd, EM_LINESCROLL, 0, -Vert 'do the scrolling
            If chkboxNoLBScroll.Value = 0 Then 'listbox scroll is on
                SendMessageBynum Thing.hwnd, LB_SETTOPINDEX, VScroll1.Value, 0 'do the scrolling
            End If
        End If
End Sub

Private Sub mnuT1SimulScroll_Click()
    Title = "Test #1"
    Msg = "Populate the text/list boxes 1st..." & vbCrLf & vbCrLf
    Msg = Msg & "Leave the two check boxes un-checked and scroll the scrollbar." & vbCrLf & vbCrLf
    Msg = Msg & "You SHOULD see all objects scroll in sync." & vbCrLf & vbCrLf
    Msg = Msg & "You will see that the List Box lags a bit." & vbCrLf
    Msg = Msg & "The list box scrolls to the correct line... but settles in by pixel (like a slot machine)." & vbCrLf
    Msg = Msg & "It literally scrolls pixel by pixel to the desired line.  Scroll slowly and you can see what I mean." & vbCrLf
    Msg = Msg & "The text boxes all 'jump' to the desired line in GoTo fashion."
    MsgBox Msg, vbOKOnly + vbInformation, Title
End Sub

Private Sub mnuT2NoListBox_Click()
    Title = "Test #2"
    Msg = "Set the check mark to Disable List Box Scrolling... and do some scrolling..." & vbCrLf & vbCrLf
    Msg = Msg & "You will see that the text boxes scroll quickly." & vbCrLf & vbCrLf
    Msg = Msg & "Leave the scrollbar in such a position that it is out of sync with the list box" & vbCrLf
    Msg = Msg & "Before doing Test 3"
    MsgBox Msg, vbOKOnly + vbInformation, Title
End Sub

Private Sub mnuT3ResyncList_Click()
    Title = "Test #3"
    Msg = "Un-check the Disable List Box Scrolling checkbox (enable listbox scrolling)" & vbCrLf & vbCrLf
    Msg = Msg & "Do some scrolling and you will see the listbox automatically re-synchronize with the scroll bar."
    MsgBox Msg, vbOKOnly + vbInformation, Title
End Sub
Private Sub mnuT4NoSync_Click()
    Title = "Test #4"
    Msg = "Set the check mark to Disable Line Synchronization" & vbCrLf & vbCrLf
    Msg = Msg & "Place your cursor in SimRTB(1)'s text area and using the DOWN ARROW," & vbCrLf
    Msg = Msg & "change the text displayed so it is displaying different lines than SimRTB(1)." & vbCrLf & vbCrLf
    Msg = Msg & "Do The Same Thing to Text1... put it on a different line than either of the others." & vbCrLf & vbCrLf
    Msg = Msg & "Do some scrolling and all the text boxes will scroll out of sync." & vbCrLf & vbCrLf
    Msg = Msg & "Leave the text boxes out of sync before doing Test #5."
    MsgBox Msg, vbOKOnly + vbInformation, Title
End Sub

Private Sub mnuT5ReSync_Click()
    Title = "Test #5"
    Msg = "UN-CHECK the check mark to Disable Line Synchronization (enables Synchronization)" & vbCrLf & vbCrLf
    Msg = Msg & "Do some scrolling and all the text boxes will automatically re-sync." & vbCrLf & vbCrLf
    Msg = Msg & "This is, I think, fairly useful code."
    
    MsgBox Msg, vbOKOnly + vbInformation, Title
End Sub

Private Sub VScroll1_Change()
    VScroll1_Scroll 'pass thru a click to VScroll1_Scroll
End Sub

Private Sub VScroll1_Scroll()

    'Dim Thing As Object
    Dim X As Integer
    'check for sync needs
    If chkboxNoSync.Value = 0 Then
        're-sync is ON... check line #'s Before scroll
        CheckLineSync
    End If
    For X = 0 To 3
        ScrollStuff SimRTB(X) 'process the Rich Text Boxes
    Next
    ScrollStuff Text1 'scroll the text1 textbox
    ScrollStuff List1 'scroll the listbox... to be honest, I doubt if this is any faster
    'than: List1.TopIndex = VScroll1.Value
    'List1.TopIndex = VScroll1.Value 'left this here for your tests
    LastScrollVal = VScroll1.Value 'store the current value for next change
End Sub
Private Sub CheckLineSync()
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '+ This sub is useful if some wiseguy has put the cursor in one of the text boxes and  +
    '+ then used the up/down arrows to move the text box text to different places.         +
    '+ It simply re-syncs all lines to the same value for simultaneous scrolling.          +
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    'I'll sync everything to the scroll bar... it could as easily be done with any of the boxes
    '... Note: LB_SETTOPINDEX in the ScrollStuff sub will automatically sync the list box to
    '... the scroll bar in the next VScroll1_Scroll or VScroll1_Change evnet... so I don't need to handle it.
    Dim FirstVis As Long, Dif As Integer, X As Integer
    For X = 0 To 3
        FirstVis = SendMessageBynum(SimRTB(X).hwnd, EM_GETFIRSTVISIBLELINE, 0, 0) 'get the topmost visible line #
        If FirstVis <> LastScrollVal Then 'can't check agains VScroll1.value... it's been moved
            Dif = LastScrollVal - FirstVis 'calculate the change needed to re-sync
            SendMessageBynum SimRTB(X).hwnd, EM_LINESCROLL, 0, Dif 'scroll there
        End If
    Next
    'do it again for the text box
    FirstVis = SendMessageBynum(Text1.hwnd, EM_GETFIRSTVISIBLELINE, 0, 0)
    If FirstVis <> LastScrollVal Then
        Dif = LastScrollVal - FirstVis
        SendMessageBynum Text1.hwnd, EM_LINESCROLL, 0, Dif
    End If
End Sub

