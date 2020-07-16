VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   14610
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   25740
   LinkTopic       =   "Form1"
   ScaleHeight     =   974
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1716
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chklocked 
      Height          =   495
      Left            =   5760
      TabIndex        =   51
      Top             =   5280
      Width           =   255
   End
   Begin VB.CommandButton cmdvertex 
      Caption         =   "ê"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   5640
      TabIndex        =   50
      Top             =   4560
      Width           =   375
   End
   Begin VB.CommandButton cmdvertex 
      Caption         =   "é"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   5640
      TabIndex        =   49
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton cmdvertex 
      Caption         =   "Expand"
      Height          =   375
      Index           =   3
      Left            =   21720
      TabIndex        =   48
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CheckBox chkdo2dpoints 
      Height          =   255
      Left            =   5280
      TabIndex        =   47
      Top             =   120
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox chkdrawconnectinglines 
      Caption         =   "Draw Connecting lines"
      Height          =   255
      Left            =   17400
      TabIndex        =   46
      Top             =   5160
      Width           =   2895
   End
   Begin VB.CheckBox chkautoadd 
      Caption         =   "Auto-add point"
      Height          =   255
      Left            =   14040
      TabIndex        =   45
      Top             =   120
      Width           =   1335
   End
   Begin VB.CheckBox chkmousewheel 
      Caption         =   "Mousewheel"
      Height          =   255
      Left            =   12600
      TabIndex        =   44
      Top             =   120
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.TextBox txtmain 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   11280
      TabIndex        =   43
      Top             =   120
      Width           =   1215
   End
   Begin VB.Timer TimerWheel 
      Interval        =   10
      Left            =   5400
      Top             =   5280
   End
   Begin VB.CommandButton cmdvertex 
      Caption         =   "Terminate"
      Height          =   375
      Index           =   2
      Left            =   17400
      TabIndex        =   41
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CheckBox chkadd 
      Caption         =   "Add Points mode"
      Height          =   255
      Left            =   9000
      TabIndex        =   40
      Top             =   120
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.CheckBox chksymmetrical 
      Caption         =   "Symmetrical model"
      Height          =   255
      Left            =   7320
      TabIndex        =   39
      Top             =   120
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.CommandButton cmdvertex 
      Caption         =   "Add Point"
      Height          =   375
      Index           =   1
      Left            =   17400
      TabIndex        =   36
      Top             =   120
      Width           =   1335
   End
   Begin VB.ListBox lstdots 
      Height          =   4155
      Left            =   17400
      TabIndex        =   35
      Top             =   480
      Width           =   1335
   End
   Begin VB.CheckBox chkis3d 
      Caption         =   "Check1"
      Height          =   255
      Left            =   240
      TabIndex        =   34
      Top             =   120
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CommandButton cmdvertex 
      Caption         =   "New Vertex"
      Height          =   375
      Index           =   0
      Left            =   15840
      TabIndex        =   33
      Top             =   120
      Width           =   1335
   End
   Begin VB.ListBox lstvertexes 
      Height          =   4545
      Left            =   15840
      TabIndex        =   32
      Top             =   480
      Width           =   1335
   End
   Begin VB.VScrollBar vsctop 
      Enabled         =   0   'False
      Height          =   4500
      Index           =   3
      Left            =   15360
      Max             =   300
      TabIndex        =   27
      Top             =   9960
      Value           =   200
      Width           =   255
   End
   Begin VB.VScrollBar vsctop 
      Enabled         =   0   'False
      Height          =   4500
      Index           =   2
      Left            =   6000
      Max             =   300
      TabIndex        =   26
      Top             =   9960
      Width           =   255
   End
   Begin VB.VScrollBar vsctop 
      Height          =   4500
      Index           =   1
      Left            =   15480
      Max             =   300
      TabIndex        =   25
      Top             =   5400
      Value           =   200
      Width           =   255
   End
   Begin VB.VScrollBar vsctop 
      Height          =   4500
      Index           =   0
      Left            =   6000
      Max             =   300
      TabIndex        =   24
      Top             =   5400
      Width           =   255
   End
   Begin VB.ListBox lstcolors 
      Height          =   2595
      Left            =   2880
      TabIndex        =   23
      Top             =   5640
      Width           =   2175
   End
   Begin VB.PictureBox imgvoyager 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   4
      Left            =   24480
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   22
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox imgvoyager 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   3
      Left            =   24480
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   21
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox imgvoyager 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   2
      Left            =   24480
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   20
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox imgvoyager 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   24480
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   19
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox imgvoyager 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   0
      Left            =   24480
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   18
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.FileListBox Filemain 
      Height          =   2625
      Left            =   600
      Pattern         =   "*.gif;*.ini;*.3d"
      TabIndex        =   17
      Top             =   5640
      Width           =   2175
   End
   Begin VB.PictureBox picvoyager 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4500
      Index           =   4
      Left            =   6360
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   596
      TabIndex        =   16
      Top             =   9960
      Width           =   8940
   End
   Begin VB.PictureBox picvoyager 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   9060
      Index           =   2
      Left            =   15840
      ScaleHeight     =   604
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   300
      TabIndex        =   14
      Top             =   5400
      Width           =   4500
   End
   Begin VB.PictureBox picvoyager 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4500
      Index           =   1
      Left            =   10920
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   300
      TabIndex        =   13
      Top             =   5400
      Width           =   4500
   End
   Begin VB.PictureBox picvoyager 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4500
      Index           =   0
      Left            =   6360
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   300
      TabIndex        =   12
      Top             =   5400
      Width           =   4500
   End
   Begin VB.PictureBox picall 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4500
      Left            =   10920
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   300
      TabIndex        =   10
      Top             =   480
      Width           =   4500
   End
   Begin VB.ListBox lstpoints 
      Height          =   5910
      Left            =   600
      TabIndex        =   8
      Top             =   8520
      Width           =   4455
   End
   Begin VB.VScrollBar vdepth 
      Height          =   4455
      Left            =   6000
      Max             =   0
      Min             =   100
      TabIndex        =   7
      Top             =   480
      Width           =   255
   End
   Begin VB.PictureBox picwireframe 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4500
      Left            =   6360
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   300
      TabIndex        =   6
      Top             =   480
      Width           =   4500
   End
   Begin VB.VScrollBar vscscale 
      Height          =   4500
      Left            =   5280
      Max             =   0
      Min             =   100
      TabIndex        =   5
      Top             =   480
      Value           =   100
      Width           =   255
   End
   Begin VB.CheckBox chklines 
      Caption         =   "Draw Lines"
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   5040
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.Timer Timermain 
      Interval        =   100
      Left            =   24600
      Top             =   120
   End
   Begin VB.CheckBox chkmain 
      Caption         =   "Auto-rotate"
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      Top             =   5040
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.HScrollBar hscmain 
      Height          =   255
      Left            =   600
      Max             =   359
      TabIndex        =   2
      Top             =   120
      Width           =   4455
   End
   Begin VB.VScrollBar VScmain 
      Height          =   4500
      Left            =   240
      Max             =   100
      TabIndex        =   1
      Top             =   480
      Value           =   25
      Width           =   255
   End
   Begin VB.PictureBox picmain 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4500
      Left            =   600
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   300
      TabIndex        =   0
      Top             =   480
      Width           =   4500
   End
   Begin VB.PictureBox picvoyager 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   9060
      Index           =   3
      Left            =   20520
      ScaleHeight     =   604
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   300
      TabIndex        =   15
      Top             =   5400
      Width           =   4500
   End
   Begin VB.Label Label1 
      Caption         =   "Lock Y:"
      Height          =   255
      Left            =   10680
      TabIndex        =   42
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblvoyager 
      Caption         =   "LCARS colors"
      Height          =   255
      Index           =   5
      Left            =   2880
      TabIndex        =   38
      Top             =   5400
      Width           =   2175
   End
   Begin VB.Label lblvoyager 
      Caption         =   "Files in this folder"
      Height          =   255
      Index           =   4
      Left            =   600
      TabIndex        =   37
      Top             =   5400
      Width           =   2175
   End
   Begin VB.Label lblvoyager 
      Caption         =   "Bottom"
      Height          =   255
      Index           =   3
      Left            =   24360
      TabIndex        =   31
      Top             =   5160
      Width           =   615
   End
   Begin VB.Label lblvoyager 
      Caption         =   "Top"
      Height          =   255
      Index           =   2
      Left            =   15840
      TabIndex        =   30
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label lblvoyager 
      Caption         =   "Back"
      Height          =   255
      Index           =   1
      Left            =   10920
      TabIndex        =   29
      Top             =   5160
      Width           =   3015
   End
   Begin VB.Label lblvoyager 
      Caption         =   "Front"
      Height          =   255
      Index           =   0
      Left            =   6360
      TabIndex        =   28
      Top             =   5160
      Width           =   3015
   End
   Begin VB.Label lblpoint 
      Height          =   255
      Left            =   600
      TabIndex        =   11
      Top             =   8280
      Width           =   4455
   End
   Begin VB.Label lbldepth 
      Caption         =   "Z = 0%"
      Height          =   255
      Left            =   6000
      TabIndex        =   9
      Top             =   120
      Width           =   1215
   End
   Begin VB.Menu Mnunew 
      Caption         =   "New"
   End
   Begin VB.Menu mnusave 
      Caption         =   "Save"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cX As Single, cY As Single, Currentfile As String

Private Sub chklines_Click()
    hscmain_Change
End Sub

Private Sub chklocked_Click()
    vsctop(0).Enabled = chklocked.Value = vbUnchecked
    vsctop(1).Enabled = vsctop(0).Enabled
    vsctop(2).Enabled = vsctop(0).Enabled
    vsctop(3).Enabled = vsctop(0).Enabled
End Sub

Private Sub chkmousewheel_Click()
    TimerWheel.Enabled = chkmousewheel.Value = vbChecked
    Debug.Print "TimerWheel.Enabled = " & TimerWheel.Enabled
End Sub

Private Sub cmdvertex_Click(Index As Integer)
    Select Case Index
        Case 0 'new vertex
            If lstcolors.ListIndex = -1 Then
                MsgBox "You must select a color first"
            Else
                lstvertexes.AddItem VertexCount & " = " & lstcolors.ListIndex
                AddVertex lstcolors.ListIndex
            End If
        Case 1 'add points
            If lstvertexes.ListIndex = -1 Then
                MsgBox "You must select a vertex"
            ElseIf lstpoints.ListIndex = -1 Then
                MsgBox "You must select a point"
            Else
                AddDotToVertex lstvertexes.ListIndex, lstpoints.ListIndex
                lstdots.AddItem lstpoints.ListIndex
            End If
        Case 2 'terminate line
            If lstvertexes.ListIndex = -1 Then
                MsgBox "You must select a vertex"
            Else
                AddDotToVertex lstvertexes.ListIndex, -1
                lstdots.AddItem "-1"
                If chksymmetrical.Value = vbChecked And lstvertexes.ListIndex < lstvertexes.ListCount And chkautoadd.Value = vbChecked Then
                    AddDotToVertex lstvertexes.ListIndex + 1, -1
                End If
            End If
        Case 3 'expand
            picvoyager(2).Width = IIf(picvoyager(2).Width = 300, 600, 300)
            picvoyager(3).Visible = (picvoyager(2).Width = 300)
        Case 4, 5 'up, down
            vdepth.Value = GetNextZ(vdepth.Value * 0.01, Index = 4) * 100
    End Select
End Sub

Private Sub Filemain_Click()
    Dim Filename As String, Extention As String, temp As Long
    If Filemain.ListIndex > -1 Then
        Filename = Filemain.List(Filemain.ListIndex)
        Extention = "." & Right(Filename, Len(Filename) - InStrRev(Filename, "."))
        If StrComp(Extention, ".3d", vbTextCompare) = 0 Or StrComp(Extention, ".ini", vbTextCompare) = 0 Then
            Filename = Left(Filename, InStr(Filename, ".") - 1)
            LoadModel chkfile(Filemain.Path, Filename & Extention)
            Extention = ".gif"
            RefreshAllLists
        Else
            Filename = Left(Filename, InStr(Filename, " ") - 1)
        End If
        Currentfile = Filename

        temp = LoadaPicture(picvoyager(0), imgvoyager(0), chkfile(Filemain.Path, Filename & " front" & Extention))
        If chklocked.Value = vbUnchecked Then
            vsctop(0).Value = picvoyager(0).Height * 0.5 - temp * 0.5
            vsctop(1).Value = picvoyager(0).Height * 0.5 + temp * 0.5
        End If
        
        LoadaPicture picvoyager(1), imgvoyager(1), chkfile(Filemain.Path, Filename & " back" & Extention)
        LoadaPicture picvoyager(2), imgvoyager(2), chkfile(Filemain.Path, Filename & " top" & Extention)
        LoadaPicture picvoyager(3), imgvoyager(3), chkfile(Filemain.Path, Filename & " bottom" & Extention)
        
        temp = LoadaPicture(picvoyager(4), imgvoyager(4), chkfile(Filemain.Path, Filename & " side" & Extention))
        If chklocked.Value = vbUnchecked Then
            vsctop(2).Value = picvoyager(4).Height * 0.5 - temp * 0.5
            vsctop(3).Value = picvoyager(4).Height * 0.5 + temp * 0.5
        End If
    End If
End Sub

Public Function LoadaPicture(Dest As PictureBox, Dest2 As PictureBox, Filename As String) As Long
    'Dim Size As Point3D
    Dest2 = LoadPicture(Filename)
    'Size = Thumbsize(Dest2.Width, Dest2.Height, Dest.Width, Dest.Height, True, False)
    
    LoadaPicture = generatethumbfromimage(Dest2, Dest, Dest.Width, Dest.Height)
    
    Dest2 = LoadPicture("")
    Dest2.Width = Dest.Width
    Dest2.Height = Dest.Height
    CopyImage Dest, Dest2
End Function

Public Sub CopyImage(Source As PictureBox, Dest As PictureBox)
    Dest.Cls
    BitBlt Dest.hdc, 0, 0, Dest.Width, Dest.Height, Source.hdc, 0, 0, vbSrcCopy
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38: cmdvertex_Click 4 'up
        Case 40: cmdvertex_Click 5 'down
        Case 80 'picture
        Case 82: chkmain.Value = IIf(chkmain.Value = vbChecked, vbUnchecked, vbChecked) 'R rotate
        Case 83: mnusave_Click 'S Save
        Case 96: cmdvertex_Click 2 '0 terminate
        Case Else: Debug.Print KeyCode
    End Select
End Sub

Private Sub Form_Load()
    Dim temp As Long
    LineColor = vbBlack
    Filemain.Path = App.Path
    AddColors
    For temp = 0 To ColorCount - 1
        lstcolors.AddItem (Colors(temp).Name)
    Next
End Sub

Private Sub hscmain_Change()
    IsClean = False
End Sub
Private Sub hscmain_Scroll()
    hscmain_Change
End Sub

Private Sub lstcolors_Click()
    If lstcolors.ListIndex > -1 And lstvertexes.ListIndex > -1 Then
        Verteces(lstvertexes.ListIndex).Color = lstcolors.ListIndex
        IsClean = False
    End If
End Sub

Private Sub lstcolors_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton And lstcolors.ListIndex > -1 Then
        LineColor = Colors(lstcolors.ListIndex).Color
    End If
End Sub

Private Sub lstdots_Click()
    If lstdots.ListIndex > -1 Then
        'lstpoints.ListIndex = lstdots.List(lstdots.ListIndex)
    End If
End Sub

Private Sub lstdots_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim temp As Long
    If (KeyCode = 46 Or KeyCode = 8) And lstdots.ListIndex > -1 And lstvertexes.ListIndex > -1 Then
        temp = lstdots.ListIndex - 1
        DeleteVertex lstvertexes.ListIndex, True, lstdots.ListIndex
        lstdots.RemoveItem lstdots.ListIndex
        lstdots.ListIndex = temp
    End If
End Sub

Private Sub lstpoints_Click()
    If lstpoints.ListIndex > -1 Then
        With Dots(lstpoints.ListIndex)
            vdepth.Value = .Z * 100
        End With
    End If
End Sub

Private Sub lstpoints_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim temp As Integer
    If (KeyCode = 46 Or KeyCode = 8) And lstpoints.ListIndex > -1 Then
        temp = lstpoints.ListIndex - 1
        RemovePoint lstpoints.ListIndex
        lstpoints.RemoveItem lstpoints.ListIndex
        lstpoints.ListIndex = temp
    End If
End Sub

Private Sub lstvertexes_Click()
    Dim temp As Long
    If lstvertexes.ListIndex > -1 Then
        lstdots.Clear
        For temp = 0 To Verteces(lstvertexes.ListIndex).PointCount - 1
            lstdots.AddItem Verteces(lstvertexes.ListIndex).Points(temp)
        Next
    End If
End Sub

Private Sub lstvertexes_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = 46 Or KeyCode = 8) And lstvertexes.ListIndex > -1 Then
        lstdots.Clear
        DeleteVertex lstvertexes.ListIndex
        lstvertexes.RemoveItem lstvertexes.ListIndex
    End If
End Sub

Private Sub Mnunew_Click()
    NewModel
    lstpoints.Clear
    lstdots.Clear
    lstvertexes.Clear
End Sub

Private Sub mnusave_Click()
    If Len(Currentfile) > 0 And DotCount > 0 Then
        SaveModel chkfile(App.Path, Currentfile & ".3d")
        If fileexists(chkfile("C:\Users\Techni\Documents\VB4A\LCAR\Files", Currentfile & ".3d")) Then
            SaveModel chkfile("C:\Users\Techni\Documents\VB4A\LCAR\Files", Currentfile & ".3d")
            MsgBox "File saved as: " & chkfile(App.Path, Currentfile & ".3d") & vbNewLine & "and: " & chkfile("C:\Users\Techni\Documents\VB4A\LCAR\Files", Currentfile & ".3d")
        Else
            MsgBox "File saved as: " & chkfile(App.Path, Currentfile & ".3d")
        End If
        Filemain.Refresh
    Else
        MsgBox "There is nothing to save"
    End If
End Sub

Private Sub picall_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseMove X, Y
End Sub


Private Sub picmain_Click()
    SavePicture picmain.Image, App.Path & "\screenshot.bmp"
    Clipboard.Clear
    Clipboard.SetText App.Path & "\screenshot.bmp"
End Sub

Private Sub picvoyager_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
End Sub

Private Sub picvoyager_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 0 'front
            If Len(txtmain) > 0 Then
                lstpoints.ListIndex = ClosestDotFront(picvoyager(0), picmain.Width * 0.5, picmain.Height * 0.5, picmain.Width * 0.5, picmain.Height * 0.5, vsctop(0).Value, vsctop(1).Value, X, Y)
                If Button = vbRightButton Then cmdvertex_Click 1
            End If
        Case 2 'top
            picwireframe_MouseDown Button, Shift, X / picvoyager(2).Width * picvoyager(0).Width, Y / picvoyager(2).Height * picvoyager(0).Height
        Case 3 'bottom
            picwireframe_MouseDown Button, Shift, picvoyager(0).Width - X, Y
    End Select
End Sub

Private Sub picvoyager_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 2 'top
            MouseMove X, Y, IIf(picvoyager(3).Visible, 0, 2)
        Case 3 'bottom
            MouseMove picvoyager(0).Width - X, Y
    End Select
End Sub



Private Sub picwireframe_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseMove X, Y
End Sub

Private Sub TimerWheel_Timer()
    Dim temp As Long
    temp = ScrollMoved(Me.hwnd)
    If temp < 0 Then temp = -15
    If temp > 0 Then temp = 15
    If temp <> 0 Then
        If temp < 0 And vdepth.Value > 0 Then vdepth.Value = vdepth.Value - 1
        If temp > 0 And vdepth.Value < 100 Then vdepth.Value = vdepth.Value + 1
    End If
End Sub

Private Sub txtmain_Click()
    chkmousewheel.Value = vbUnchecked
End Sub

Private Sub txtmain_GotFocus()
    chkmousewheel.Value = vbUnchecked
End Sub

Private Sub txtmain_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then chkmousewheel.Value = vbChecked
End Sub

Private Sub txtmain_LostFocus()
    chkmousewheel.Value = vbChecked
End Sub

Private Sub vdepth_Change()
    hscmain_Change
End Sub
Private Sub vdepth_Scroll()
    hscmain_Change
    lbldepth = "Z = " & vdepth.Value & "%"
End Sub
Private Sub VScmain_Change()
    hscmain_Change
End Sub
Private Sub VScmain_Scroll()
    hscmain_Change
End Sub

Public Sub MouseMove(X As Single, Y As Single, Optional Index As Integer)
    Dim WasChecked As Boolean, temp As Point3D '
    If Len(txtmain) > 0 And IsNumeric(txtmain) Then Y = txtmain
    Me.Caption = X & ", " & Y
    WasChecked = chkmain.Value = vbChecked
    cX = X
    cY = Y
    If Index > 0 Then
        cX = cX / picvoyager(Index).Width * picvoyager(0).Width
        cY = cY / picvoyager(Index).Height * picvoyager(0).Height
    End If
    If WasChecked Then chkmain.Value = vbUnchecked
    IsClean = False
    Timermain_Timer
    If WasChecked Then chkmain.Value = vbChecked
    
    temp.Radius = Distance(picwireframe.Width * 0.5, picwireframe.Height * 0.5, X, Y) / (picwireframe.Width * 0.5)
    temp.Angle = GetAngle(picwireframe.Width * 0.5, picwireframe.Height * 0.5, X, Y)
    temp.Z = vdepth.Value * 0.01
    lblpoint.Caption = "Radius: " & FormatPercent(temp.Radius, 2) & "    Angle: " & temp.Angle & "°  Z: " & FormatPercent(temp.Z, 0)
End Sub






Private Sub Timermain_Timer()
    Dim temp As Long, X As Long, Y As Long
    If chkmain.Value = vbChecked Then hscmain.Value = (hscmain.Value + 1) Mod 360
    
    If Not IsClean Then
        picmain.Cls
        If chklines.Value = vbChecked Then
            temp = VScmain.Value * 0.01 * picmain.Height
            DrawLine picmain, 0, temp, picmain.Width, temp, LineColor
            DrawOval picmain, 0, temp, picmain.Width - 1, picmain.Height - temp - 1, LineColor

            X = findXYAngle(picmain.Width * 0.5, temp + (picmain.Height - temp) * 0.5, picmain.Width * 0.5, hscmain.Value, True)
            Y = findXYAngle(picmain.Width * 0.5, temp + (picmain.Height - temp) * 0.5, (picmain.Height - temp) * 0.5, hscmain.Value, False)
            DrawLine picmain, picmain.Width * 0.5, temp + (picmain.Height - temp) * 0.5, X, Y, LineColor
            DrawLine picmain, X, Y, X, Y - vscscale.Value, LineColor
        End If
        DrawDots picmain, hscmain.Value, picmain.Width * 0.5, temp + (picmain.Height - temp) * 0.5, picmain.Width * 0.5, (picmain.Height - temp) * 0.5, picmain.Height * (vscscale.Value * 0.01), -1, IIf(chkdo2dpoints.Value = vbChecked, -1, LineColor), lstpoints.ListIndex, chkis3d.Value = vbChecked
        If chkis3d.Value = vbChecked Then DrawVerteces picmain
        
        DrawWireframe picwireframe, vdepth.Value * 0.01, LineColor
        DrawWireframe picall, -1, LineColor
        
        temp = GetValue(vsctop(0).Value, vsctop(1).Value, vdepth.Value * 0.01, True)
       
        'front
        CopyImage imgvoyager(0), picvoyager(0)
        DrawLine picvoyager(0), 0, vsctop(0).Value, picvoyager(0).Width, vsctop(0).Value, vbRed 'top
        DrawLine picvoyager(0), 0, vsctop(1).Value, picvoyager(0).Width, vsctop(1).Value, vbRed 'bottom
        DrawLine picvoyager(0), 0, temp, picvoyager(0).Width, temp, LineColor 'Z
        DrawLine picvoyager(0), cX, 0, cX, picvoyager(0).Height, LineColor 'X
        If Len(txtmain) > 0 Then DrawDotsFront picvoyager(0), picmain.Width * 0.5, picmain.Height * 0.5, picmain.Width * 0.5, picmain.Height * 0.5, vsctop(0).Value, vsctop(1).Value, LineColor, lstpoints.ListIndex
        picvoyager(0).Refresh
        
        'back
        CopyImage imgvoyager(1), picvoyager(1)
        DrawLine picvoyager(1), 0, vsctop(0).Value, picvoyager(0).Width, vsctop(0).Value, vbRed 'top
        DrawLine picvoyager(1), 0, vsctop(1).Value, picvoyager(0).Width, vsctop(1).Value, vbRed 'bottom
        DrawLine picvoyager(1), 0, temp, picvoyager(0).Width, temp, LineColor 'Z
        DrawLine picvoyager(1), cX, 0, cX, picvoyager(0).Height, LineColor 'inverted X
        picvoyager(1).Refresh
        
        'top
        'If picvoyager(3).Visible Then
            temp = cY / picvoyager(0).Height * picvoyager(2).Height
            X = cX / picvoyager(0).Width * picvoyager(2).Height
            X = (picvoyager(2).Width * 0.5) - (picvoyager(2).Height * 0.5) + X
        'Else
        '    temp = cY / picvoyager(0).Width * picvoyager(2).Height
        '    X = cX / picvoyager(0).Width * picvoyager(2).Height
        '    X = (picvoyager(2).Width * 0.5) - (picvoyager(2).Height * 0.5) + X
       ' End If
        
        CopyImage imgvoyager(2), picvoyager(2)
        DrawLine picvoyager(2), 0, temp, picvoyager(2).Width, temp, LineColor 'y
        DrawLine picvoyager(2), X, 0, X, picvoyager(2).Height, LineColor 'x
        DrawWireframe picvoyager(2), vdepth.Value * 0.01, LineColor, False, chkdrawconnectinglines.Value = vbChecked
        picvoyager(2).Refresh
        
        'bottom
        If picvoyager(3).Visible Then
            CopyImage imgvoyager(3), picvoyager(3)
            DrawLine picvoyager(3), 0, temp, picvoyager(3).Width, temp, LineColor 'y
            DrawLine picvoyager(3), X, 0, X, picvoyager(3).Height, LineColor 'x MUST CORRECT FOR NEW SCALE!
            DrawWireframe picvoyager(3), vdepth.Value * 0.01, LineColor, False, chkdrawconnectinglines.Value = vbChecked
            picvoyager(3).Refresh
        End If
        
        'side
        temp = GetValue(vsctop(2).Value, vsctop(3).Value, vdepth.Value * 0.01, True)
        CopyImage imgvoyager(4), picvoyager(4)
        DrawLine picvoyager(4), 0, vsctop(2).Value, picvoyager(4).Width, vsctop(2).Value, vbRed 'top
        DrawLine picvoyager(4), 0, vsctop(3).Value, picvoyager(4).Width, vsctop(3).Value, vbRed 'bottom
        DrawLine picvoyager(4), 0, temp, picvoyager(4).Width, temp, LineColor 'Z
        
        temp = cY / picvoyager(0).Height * picvoyager(4).Width
        DrawLine picvoyager(4), temp, 0, temp, picvoyager(0).Height, LineColor
        picvoyager(4).Refresh
        
        IsClean = True
    End If
End Sub

Sub DrawWireframe(Dest As PictureBox, Depth As Single, Optional Color As OLE_COLOR = vbBlack, Optional DoClear As Boolean = True, Optional DrawLines As Boolean = True)
    Dim Radius As Long
    Radius = Max(Dest.Height, Dest.Width) * 0.5
    If DoClear Then
        Dest.Cls
        DrawLine Dest, 0, cY, Dest.Width, cY, vbRed
        DrawLine Dest, cX, 0, cX, Dest.Height, vbRed
        DrawLine Dest, Dest.Width * 0.5, Dest.Height * 0.5, Dest.Width * 0.5, 0, Color
    End If
    DrawCircle Dest, Dest.Width * 0.5, Dest.Height * 0.5, Dest.Height * 0.5 - 1, Color, Color
    DrawDots Dest, 0, Dest.Width * 0.5, Dest.Height * 0.5, Radius, Radius, 0, Depth, Color, lstpoints.ListIndex, , DrawLines
End Sub



Private Sub picall_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picwireframe_MouseDown Button, Shift, X, Y
End Sub

Private Sub picwireframe_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim temp As Point3D, temp2 As Long, tempDistance As Long, Closest As Long, eDistance As Long, Radius As Long
    If Len(txtmain) > 0 And IsNumeric(txtmain) Then Y = txtmain
    temp.Radius = Distance(picwireframe.Width * 0.5, picwireframe.Height * 0.5, X, Y) / (picwireframe.Width * 0.5)
    temp.Angle = GetAngle(picwireframe.Width * 0.5, picwireframe.Height * 0.5, X, Y)
    temp.Z = vdepth.Value * 0.01
    If chkadd.Value = vbChecked Then
        AddPoint temp
        lstpoints.AddItem (DotCount - 1) & " = Radius: " & FormatPercent(temp.Radius, 2) & "    Angle: " & temp.Angle & "°  Z: " & FormatPercent(temp.Z, 0)
        lstpoints.ListIndex = lstpoints.ListCount - 1
        If chkautoadd.Value = vbChecked And lstvertexes.ListIndex > -1 Then
            If lstvertexes.ListIndex + Shift >= VertexCount Then
                cmdvertex_Click 0 'add vertex
            End If
            AddDotToVertex lstvertexes.ListIndex + Shift, lstpoints.ListIndex
            If Shift = 0 Then lstdots.AddItem lstpoints.ListIndex
        End If
        If Shift = 0 And chksymmetrical.Value = vbChecked Then
            If X < (picwireframe.Width * 0.5 - 5) Or X > (picwireframe.Width * 0.5 + 5) Then 'ignore middle dots
                picwireframe_MouseDown Button, 1, picwireframe.Width - X, Y
            End If
        End If
    Else
        Closest = -1
        Radius = picwireframe.Width * 0.5
        For temp2 = 0 To DotCount - 1
            With Dots(temp2)
                If .Z = temp.Z Then
                    tempDistance = Distance(X, Y, findXYAngle(Radius, Radius, .Radius * Radius, .Angle, True), findXYAngle(Radius, Radius, .Radius * Radius, .Angle, False))
                    Debug.Print
                    If Closest = -1 Or eDistance > tempDistance Then
                        Closest = temp2
                        eDistance = tempDistance
                    End If
                End If
            End With
        Next
        If Closest > -1 Then
            lstpoints.ListIndex = Closest
            If Button = vbRightButton Then
                cmdvertex_Click 1
            End If
        End If
    End If
End Sub

Private Sub picwireframe_KeyDown(KeyCode As Integer, Shift As Integer)
    Debug.Print "picwireframe_KeyDown: " & KeyCode
    If lstpoints.ListIndex > -1 Then
        IsLocked = True
        With Dots(lstpoints.ListIndex)
            Select Case KeyCode
                Case 37 'Left
                    .Angle = .Angle - 1
                    If .Angle < 0 Then .Angle = .Angle + 360
                Case 39 'Right
                    .Angle = (.Angle + 1) Mod 360
                Case 38 'up
                    .Radius = .Radius + 0.01
                    If .Radius > 1 Then .Radius = 1
                Case 40 'down
                    .Radius = .Radius - 0.01
                    If .Radius < 0 Then .Radius = 0
            End Select
        End With
        IsLocked = False
        IsClean = False
    End If
End Sub

Public Sub RefreshAllLists()
    Dim temp As Long
    lstpoints.Clear
    For temp = 0 To DotCount - 1
        With Dots(temp)
            lstpoints.AddItem temp & " = Radius: " & FormatPercent(.Radius, 2) & "    Angle: " & .Angle & "°  Z: " & FormatPercent(.Z, 0)
        End With
    Next
    If DotCount > 0 Then lstpoints.ListIndex = 0
    
    lstvertexes.Clear
    For temp = 0 To VertexCount - 1
        With Verteces(temp)
            lstvertexes.AddItem temp & " = " & .Color
        End With
    Next
    If VertexCount > 0 Then lstvertexes.ListIndex = 0
End Sub

Private Sub vscscale_Change()
    IsClean = False
    hscmain_Change
End Sub
