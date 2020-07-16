Attribute VB_Name = "Module1"
Option Explicit

Public IsClean As Boolean
Public Const PI As Double = 3.14159265358979
Public Const TransColor As Long = vbWhite 'rgb(255,0,128)
Public LineColor As Long

Type Point3D
    Angle As Long
    Radius As Single
    Z As Single
    cX As Single
    cY As Single
End Type
Type Vertex
    Color As OLE_COLOR
    Points() As Long
    PointCount As Long
End Type
Type LCARColor
    Color As OLE_COLOR
    Name As String
End Type

Public Colors() As LCARColor, ColorCount As Long
Public Dots() As Point3D, DotCount As Long
Public Verteces() As Vertex, VertexCount As Long
Public IsLocked As Boolean

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal hStretchMode As Long) As Long
Public Const STRETCHMODE = vbPaletteModeNone

Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Type msg
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type
Private Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Private Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As msg) As Long

'colors
Public Sub AddLCARcolor(Name As String, R As Byte, G As Byte, B As Byte)
    Dim temp As LCARColor
    temp.Name = Name
    temp.Color = RGB(R, G, B)
    ColorCount = ColorCount + 1
    ReDim Preserve Colors(ColorCount)
    Colors(ColorCount - 1) = temp
End Sub
Public Sub AddColors()
    If ColorCount = 0 Then
        AddLCARcolor "Black", 0, 0, 0
        AddLCARcolor "Dark Orange", 215, 107, 0
        AddLCARcolor "Orange", 253, 153, 0
        AddLCARcolor "Light Orange", 255, 255, 0
        AddLCARcolor "Purple", 255, 0, 255
        AddLCARcolor "Light Purple", 204, 153, 204
        AddLCARcolor "Light Blue", 153, 153, 204
        AddLCARcolor "Red", 204, 102, 102
        AddLCARcolor "Yellow", 255, 255, 0
        AddLCARcolor "Dark Blue", 153, 153, 255
        AddLCARcolor "Dark Yellow", 255, 153, 102
        AddLCARcolor "Dark Purple", 204, 102, 153
        AddLCARcolor "White", 128, 128, 128
        AddLCARcolor "Red Alert", 204, 102, 102
        AddLCARcolor "Light Green", 152, 255, 102
        AddLCARcolor "Green", 6, 138, 3
        AddLCARcolor "Lighter Blue", 153, 205, 255
        AddLCARcolor "Blue", 0, 0, 254
        AddLCARcolor "Turq", 76, 232, 185
        AddLCARcolor "Grey", 128, 128, 128
        AddLCARcolor "LBlue", 158, 193, 225
        AddLCARcolor "Light Yellow", 225, 239, 160
        AddLCARcolor "BORG", 0, 120, 0
        AddLCARcolor "Chrono", 15, 65, 124
    End If
End Sub








Public Function GetNextZ(CurrentZ As Single, Up As Boolean) As Single
    Dim temp As Long, NextZ As Single
    NextZ = IIf(Up, 1, 0)
    For temp = 0 To DotCount - 1
        With Dots(temp)
            If Up Then
                If Dots(temp).Z > CurrentZ And NextZ > Dots(temp).Z Then NextZ = Dots(temp).Z
            Else
                If Dots(temp).Z < CurrentZ And NextZ < Dots(temp).Z Then NextZ = Dots(temp).Z
            End If
        End With
    Next
    GetNextZ = NextZ
End Function


'3d model data api
Public Sub NewModel(Optional EraseDots As Boolean = True, Optional EraseVerteces As Boolean = True)
    IsLocked = True
    If EraseDots Then
        DotCount = 0
        ReDim Dots(0)
    End If
    If EraseVerteces Then
        VertexCount = 0
        ReDim Verteces(0)
    End If
    IsClean = False
    IsLocked = False
End Sub
Public Sub AddPoint(Point As Point3D)
    IsLocked = True
    DotCount = DotCount + 1
    ReDim Preserve Dots(DotCount)
    Dots(DotCount - 1) = Point
    IsLocked = False
    IsClean = False
End Sub
Public Function RemovePoint(Index As Long) As Long
    Dim temp As Long, temp2 As Long, temp3 As Long, Ret As Long
    Ret = -1
    If Index > -1 Then
        IsLocked = True
        For temp = Index To DotCount - 2
            Dots(temp) = Dots(temp + 1)
        Next
        DotCount = DotCount - 1
        If DotCount > 0 Then
            ReDim Preserve Dots(DotCount)
        Else
            ReDim Dots(0)
        End If
        'remove from verteces
        For temp = 0 To VertexCount - 1
            For temp2 = Verteces(temp).PointCount - 1 To 0 Step -1
                If Verteces(temp).Points(temp2) = Index Then
                    For temp3 = temp2 To Verteces(temp).PointCount - 2
                        Verteces(temp).Points(temp3) = Verteces(temp).Points(temp3 + 1)
                        Ret = Ret + 1
                    Next
                    Verteces(temp).PointCount = Verteces(temp).PointCount - 1
                    If Verteces(temp).PointCount = 0 Then
                        ReDim Verteces(temp).Points(0)
                    Else
                        ReDim Preserve Verteces(temp).Points(Verteces(temp).PointCount)
                    End If
                End If
            Next
        Next
        IsClean = False
        IsLocked = False
    End If
    RemovePoint = Ret
End Function

Public Sub AddVertexFromString(Text As String)
    Dim tempstr() As String, temp As Long, VertexID As Long
    tempstr = Split(Text, ",") 'color,list of dots
    VertexID = AddVertex(Val(tempstr(0)))
    For temp = 1 To UBound(tempstr)
        AddDotToVertex VertexID, Val(tempstr(temp))
    Next
End Sub

Public Function AddVertex(Color As Long) As Long 'Public Verteces() As Vertex, VertexCount As Long
    IsLocked = True
    AddVertex = VertexCount
    VertexCount = VertexCount + 1
    ReDim Preserve Verteces(VertexCount)
    Verteces(VertexCount - 1).Color = Color
    IsLocked = False
    IsClean = False
End Function
Public Sub AddDotToVertex(VertexID As Long, DotID As Long)
    IsLocked = True
    Verteces(VertexID).PointCount = Verteces(VertexID).PointCount + 1
    ReDim Preserve Verteces(VertexID).Points(Verteces(VertexID).PointCount)
    Verteces(VertexID).Points(Verteces(VertexID).PointCount - 1) = DotID
    IsLocked = False
    IsClean = False
End Sub

Public Sub DeleteVertex(VertexID As Long, Optional JustTheDots As Boolean, Optional DotIndex As Long = -1)
    Dim temp As Long
    IsLocked = True
    If JustTheDots Then
        If DotIndex = -1 Then 'delete all dots
            Verteces(VertexID).PointCount = 0
            ReDim Verteces(VertexID).Points(0)
        Else 'delete 1 dot
            For temp = DotIndex To Verteces(VertexID).PointCount - 2
                Verteces(VertexID).Points(temp) = Verteces(VertexID).Points(temp + 1)
            Next
            Verteces(VertexID).PointCount = Verteces(VertexID).PointCount - 1
            If Verteces(VertexID).PointCount = 0 Then
                ReDim Verteces(VertexID).Points(0)
            Else
                ReDim Preserve Verteces(VertexID).Points(Verteces(VertexID).PointCount)
            End If
        End If
    Else 'delete a vertex
        For temp = VertexID To VertexCount - 2
            Verteces(temp) = Verteces(temp + 1)
        Next
        VertexCount = VertexCount - 1
        If VertexCount = 0 Then
            ReDim Verteces(0)
        Else
            ReDim Preserve Verteces(VertexCount)
        End If
    End If
    IsLocked = False
    IsClean = False
End Sub











'file handling
Public Function chkfile(Path As String, File As String) As String
    chkfile = Replace(Path & "\" & File, "\\", "\")
End Function
Public Function fileexists(Filename As String) As Boolean
    On Error Resume Next
    If Len(Filename) > 0 Then fileexists = Len(Dir(Filename, vbNormal + vbHidden + vbSystem)) > 0
End Function
Public Sub SaveModel(Filename As String)
    On Error Resume Next
    Dim tempfile As Integer, temp As Long
    tempfile = FreeFile
    IsLocked = True
    Open Filename For Output As tempfile
        Print #tempfile, "[dots]"
        Print #tempfile, "dots=" & DotCount
        For temp = 0 To DotCount - 1
            With Dots(temp)
                Print #tempfile, temp & "=" & .Angle & "," & .Radius & "," & .Z
            End With
        Next
        Print #tempfile, "[lines]"
        Print #tempfile, "lines=" & VertexCount
        For temp = 0 To VertexCount - 1
            Print #tempfile, temp & "=" & GetVertex(temp)
        Next
    Close tempfile
    IsLocked = False
End Sub
Public Function GetVertex(Index As Long) As String
    Dim tempstr As String, temp As Long
    With Verteces(Index)
        tempstr = .Color
        For temp = 0 To .PointCount - 1
            tempstr = tempstr & "," & .Points(temp)
        Next
    End With
    GetVertex = tempstr
End Function
Public Function LoadModel(Filename As String) As Boolean
    On Error Resume Next
    Dim tempstr() As String, temp As Long, CurrentSection As String, Key As String, Value As String
    tempstr = Split(LoadFile(Filename), vbNewLine)
    NewModel
    IsLocked = True
    For temp = 0 To UBound(tempstr)
        tempstr(temp) = Trim(tempstr(temp))
        If Left(tempstr(temp), 1) = "[" And Right(tempstr(temp), 1) = "]" Then
            CurrentSection = Mid(tempstr(temp), 2, Len(tempstr(temp)) - 2)
        ElseIf Len(tempstr(temp)) > 0 And Left(tempstr(temp), 1) <> "#" Then
            Key = Left(tempstr(temp), InStr(tempstr(temp), "=") - 1)
            Value = Right(tempstr(temp), Len(tempstr(temp)) - InStr(tempstr(temp), "="))
            If IsNumeric(Key) Then
                If CurrentSection = "dots" Then
                    AddPoint MakePointFromValue(Value)
                ElseIf CurrentSection = "lines" Then
                    AddVertexFromString Value
                End If
            End If
        End If
    Next
    IsLocked = False
    IsClean = False
    LoadModel = True
End Function
Public Function LoadFile(Filename As String) As String
    On Error Resume Next
    If FileLen(Filename) = 0 Then Exit Function
    Dim temp As Long, tempstr As String, tempstr2 As String
    temp = FreeFile
    If Dir(Filename) <> Filename Then
        Open Filename For Input As temp
            Do Until EOF(temp)
                Line Input #temp, tempstr
                If tempstr2 <> Empty Then tempstr2 = tempstr2 & vbNewLine
                tempstr2 = tempstr2 & tempstr
                DoEvents
            Loop
            LoadFile = tempstr2
        Close temp
    End If
End Function
Public Function MakePointFromValue(tempstr As String) As Point3D
    Dim tempstr2() As String, Point As Point3D
    tempstr2 = Split(tempstr, ",") '.Angle & "," & .Radius & "," & .Z
    Point.Angle = tempstr2(0)
    Point.Radius = tempstr2(1)
    Point.Z = tempstr2(2)
    MakePointFromValue = Point
End Function










'Graphics
Public Sub DrawDots(Dest As PictureBox, Angle As Long, CenterX As Long, CenterY As Long, RadiusX As Long, RadiusY As Long, FullZ As Long, Optional Depth As Single = -1, Optional Color As OLE_COLOR = vbBlack, Optional SelectedDot As Long = -1, Optional DoCache As Boolean, Optional DrawLines As Boolean = True)
    Dim temp As Long, X As Long, Y As Long, A As Long, LineSize As Long, CurrentY1 As Long, CurrentY2 As Long, Color2 As Long
    If Not IsLocked Then
        If Depth > -1 Then
            LineSize = Dest.TextHeight("9")
            CurrentY1 = LineSize
            CurrentY2 = LineSize
        End If
        For temp = 0 To DotCount - 1
            With Dots(temp)
                If Depth = -1 Or Depth = .Z Then
                    A = (.Angle + Angle) Mod 360
                    X = findXYAngle(CenterX, CenterY, .Radius * RadiusX, A, True)
                    Y = findXYAngle(CenterX, CenterY, .Radius * RadiusY, A, False)
                    If FullZ > 0 Then Y = Y - (FullZ * .Z)
                    If DoCache Then
                        Dots(temp).cX = X '/ Dest.Width
                        Dots(temp).cY = Y '/ Dest.Height
                    End If
                    If Color > -1 Then
                        Color2 = IIf(temp = SelectedDot, vbRed, Color)
                        DrawPoint Dest, X, Y, Color
                    End If
                    If Depth > -1 And DrawLines Then
                        DrawDotLabel Dest, X, Y, CurrentY1, CurrentY2, Color2, temp, LineSize
                        'If X < Dest.Width * 0.5 Then
                        '    DrawLine Dest, 0, CurrentY1, 50, CurrentY1, Color2
                        '    DrawLine Dest, 50, CurrentY1, X, Y, Color2
                        '    DrawText Dest, 0, CurrentY1 - LineSize, Temp & "", , Color2
                        '    CurrentY1 = CurrentY1 + LineSize
                        'Else
                        '    DrawLine Dest, Dest.Width, CurrentY2, Dest.Width - 50, CurrentY2, Color2
                        '    DrawLine Dest, Dest.Width - 50, CurrentY2, X, Y, Color2
                        '    DrawText Dest, Dest.Width - Dest.TextWidth(Temp & ""), CurrentY2 - LineSize, Temp & "", , Color2
                        '    CurrentY2 = CurrentY2 + LineSize
                        'End If
                    End If
                End If
            End With
        Next
    End If
End Sub
Public Sub DrawDotLabel(Dest As PictureBox, X As Long, Y As Long, ByRef CurrentY1 As Long, ByRef CurrentY2 As Long, Color2 As Long, temp As Long, LineSize As Long)
    If X < Dest.Width * 0.5 Then
        DrawLine Dest, 0, CurrentY1, 50, CurrentY1, Color2
        DrawLine Dest, 50, CurrentY1, X, Y, Color2
        DrawText Dest, 0, CurrentY1 - LineSize, temp & "", , Color2
        CurrentY1 = CurrentY1 + LineSize
    Else
        DrawLine Dest, Dest.Width, CurrentY2, Dest.Width - 50, CurrentY2, Color2
        DrawLine Dest, Dest.Width - 50, CurrentY2, X, Y, Color2
        DrawText Dest, Dest.Width - Dest.TextWidth(temp & ""), CurrentY2 - LineSize, temp & "", , Color2
        CurrentY2 = CurrentY2 + LineSize
    End If
End Sub
Public Sub DrawDotsFront(Dest As PictureBox, CenterX As Long, CenterY As Long, RadiusX As Long, RadiusY As Long, Scale1 As Long, Scale2 As Long, Optional Color As OLE_COLOR = vbBlack, Optional SelectedDot = -1)
    Dim temp As Long, X As Long, Y As Long, Color2 As Long, LineSize As Long, CurrentY1 As Long, CurrentY2 As Long
    If Not (IsLocked) Then
        LineSize = Dest.TextHeight("9")
        CurrentY1 = LineSize
        CurrentY2 = LineSize
        For temp = 0 To DotCount - 1
            With Dots(temp)
                X = findXYAngle(CenterX, CenterY, .Radius * RadiusX, .Angle, True)
                Y = GetValue(Scale1, Scale2, .Z, True)
                Color2 = IIf(temp = SelectedDot, vbRed, Color)
                DrawPoint Dest, X, Y, Color2
                'DrawDotLabel Dest, X, Y, CurrentY1, CurrentY2, Color2, Temp, LineSize
            End With
        Next
    End If
End Sub
Public Function ClosestDotFront(Dest As PictureBox, CenterX As Long, CenterY As Long, RadiusX As Long, RadiusY As Long, Scale1 As Long, Scale2 As Long, X As Single, Y As Single) As Long
    Dim temp As Long, X2 As Long, Y2 As Long, eDistance As Long, tempDistance As Long
    eDistance = -1
    ClosestDotFront = -1
    If Not (IsLocked) Then
        For temp = 0 To DotCount - 1
            With Dots(temp)
                X2 = findXYAngle(CenterX, CenterY, .Radius * RadiusX, .Angle, True)
                Y2 = GetValue(Scale1, Scale2, .Z, True)
                tempDistance = Distance(X, Y, CSng(X2), CSng(Y2))
                If tempDistance < eDistance Or eDistance = -1 Then
                    ClosestDotFront = temp
                    eDistance = tempDistance
                End If
            End With
        Next
    End If
End Function

Public Sub DrawVerteces(Dest As PictureBox)
    Dim temp As Long, temp2 As Long, Color As Long, LastVertex As Long
    LastVertex = -1
    If Not IsLocked Then
        For temp = 0 To VertexCount - 1
            With Verteces(temp)
                If .PointCount > 1 Then
                    Color = Colors(.Color).Color
                    For temp2 = 0 To .PointCount - 2
                        If .Points(temp2 + 1) <> LastVertex Then
                            ConnectTheDots Dest, Color, .Points(temp2), .Points(temp2 + 1)
                        End If
                        LastVertex = .Points(temp2)
                    Next
                    ConnectTheDots Dest, Color, .Points(.PointCount - 1), .Points(0)
                End If
            End With
        Next
    End If
End Sub
Public Sub ConnectTheDots(Dest As PictureBox, Color As OLE_COLOR, DotID1 As Long, DotID2 As Long)
    If DotID1 > -1 And DotID2 > -1 Then
        DrawLine Dest, Dots(DotID1).cX, Dots(DotID1).cY, Dots(DotID2).cX, Dots(DotID2).cY, Color
    End If
End Sub

Public Sub DrawPoint(Dest As PictureBox, X As Long, Y As Long, Color As OLE_COLOR)
    Const s As Integer = 3
    Dest.Point X, Y
    
    DrawLine Dest, X - s, Y, X + s + 1, Y, Color
    DrawLine Dest, X, Y - s, X, Y + s + 1, Color
End Sub
Public Sub DrawLine(Dest As PictureBox, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, Color As OLE_COLOR)
    Dest.Line (X1, Y1)-(X2, Y2), Color
End Sub
Public Sub DrawSquare(Dest As PictureBox, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, Color As OLE_COLOR, Optional FillColor As OLE_COLOR = -1)
    Dest.DrawWidth = 1
    If FillColor > -1 Then Dest.FillColor = FillColor
    Dest.Line (X, Y)-(X + Width, Y + Height), Color, B
End Sub
Public Sub DrawCircle(Dest As PictureBox, ByVal X As Long, ByVal Y As Long, Radius As Long, Optional EdgeColor As Long = vbBlack, Optional FillColor As Long = vbBlack)
    If FillColor = EdgeColor Then
        Dest.Fillstyle = vbFSTransparent
    Else
        Dest.Fillstyle = vbSolid
        Dest.FillColor = FillColor
    End If
    Dest.Circle (X, Y), Radius, FillColor
End Sub
Public Sub DrawOval(Dest As PictureBox, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, Color As OLE_COLOR, Optional FillColor As Long = vbGreen, Optional Fillstyle As FillStyleConstants = vbFSTransparent)
    Dest.Fillstyle = Fillstyle
    Dest.FillColor = FillColor
    Dest.Circle (X + Width * 0.5, Y + Height * 0.5), Width * 0.5, Color, , , Height / Width
End Sub
 
Public Sub DrawText(Dest As PictureBox, X As Long, Y As Long, Text As String, Optional size As Long, Optional Color As OLE_COLOR = vbBlack)
    Dest.CurrentX = X
    Dest.CurrentY = Y
    Dest.ForeColor = Color
    If size > 0 Then Dest.FontSize = size
    Dest.Print Text
End Sub









'Trig
Public Function DegToRad(ByVal Deg As Double) As Double
    DegToRad = (Deg / 180) * PI
End Function
Public Function RadToDeg(ByVal Rad As Double) As Double
    RadToDeg = Rad * (180 / PI)
End Function

Public Function Angle(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single) As Double
    On Error Resume Next
    Angle = Atn((Y2 - Y1) / (X1 - X2))
End Function

Public Function CorrectAngle(ByVal Angle As Long) As Long
    Do While Angle < 0
        Angle = Angle + 360
    Loop
    CorrectAngle = Angle Mod 360
End Function

Public Function findXY(X As Single, Y As Single, Distance As Single, Angle As Double, Optional IsX As Boolean = True) As Single
    If IsX Then findXY = X + Sin(Angle) * Distance Else findXY = Y + Cos(Angle) * Distance
End Function

Public Function findXYAngle(X As Long, Y As Long, Distance As Long, Angle As Long, Optional IsX As Boolean = True) As Long
    findXYAngle = findXY(CLng(X), CLng(Y), CLng(Distance), DegToRad(CorrectAngle(180 - Angle)), IsX)
End Function

Public Function GetAngle(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single) As Long
    GetAngle = CorrectAngle(AngleBySection(X1, Y1, X2, Y2, RadToDeg(Angle(X1, Y1, X2, Y2))) - 180)
End Function


Public Function AngleBySection(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, ByVal Angle As Long) As Double
    Angle = Abs(Angle)
    AngleBySection = 90 'Corrected
    If X1 < X2 Then 'the point is at the left of Center
        If Y1 = Y2 Then
            AngleBySection = 270 'Corrected
        ElseIf Y1 < Y2 Then
            If 270 + Angle = 360 Then
                AngleBySection = 0 'Corrected
            Else
                AngleBySection = 270 + Angle 'Corrected
            End If
        ElseIf Y1 > Y2 Then
            AngleBySection = 270 - Angle 'Corrected
        End If
    Else
    
        If X1 > X2 Then 'the point is at the right of Center
            If Y1 > Y2 Then
                AngleBySection = 90 + Angle 'Corrected
            ElseIf Y1 < Y2 Then
                AngleBySection = 90 - Angle 'Corrected
            End If
        Else
    
            If X1 = X2 Then
                If Y1 < Y2 Then
                    AngleBySection = 0 'Corrected
                ElseIf Y1 > Y2 Then
                    AngleBySection = 180 'Corrected
                End If
            End If
    
        End If

    End If
End Function

Public Function Distance(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single) As Single
    On Error Resume Next
    If Y2 - Y1 = 0 Then Distance = Abs(X2 - X1): Exit Function
    If X2 - X1 = 0 Then Distance = Abs(Y2 - Y1): Exit Function
    Distance = Abs(Y2 - Y1) / Sin(Atn(Abs(Y2 - Y1) / Abs(X2 - X1)))
End Function

Public Function AngleDifference(Angle1 As Long, Angle2 As Long, Optional Absolute As Boolean) As Long
    Dim temp As Long
    temp = Angle2 - Angle1
    If temp > 180 Then temp = -360 + temp
    If Absolute Then temp = Abs(Absolute)
    AngleDifference = temp
End Function

Public Function TestAngle(X As Long, Y As Long, Distance As Long, Angle As Long) As String
    TestAngle = findXYAngle(X, Y, Distance, Angle, True) & ", " & findXYAngle(X, Y, Distance, Angle, False)
End Function
'Public Function GetPoint3D(Dest As PictureBox, X As Single, Y As Single)
    'Dim temp As Point3D 'does not appear to be used
    'temp.Radius = Distance(Dest.Width * 0.5, Dest.Height * 0.5, X, Y) / (Dest.Width * 0.5)
   ' temp.Angle = GetAngle(Dest.Width * 0.5, Dest.Height * 0.5, X, Y)
    'temp.Z = vdepth.Value * 0.01'vdepth is not found
    'GetPoint3D = temp
'End Function













'thumbnail generation
Public Function Thumbsize(ByRef PicWidth As Long, ByRef PicHeight As Long, ByRef thumbwidth As Long, ByRef thumbheight As Long, Optional ForceFit As Boolean, Optional ForceFull As Boolean) As Point3D
    Dim temp As Point3D
    If ForceFit Then
        If PicHeight < thumbheight Then
            PicWidth = PicWidth * thumbheight / PicHeight
            PicHeight = thumbheight
        End If
    End If
    If PicWidth > thumbwidth Then
        PicHeight = Round(PicHeight / (PicWidth / thumbwidth), 0)
        PicWidth = thumbwidth
    End If
    If PicHeight > thumbheight Then
        PicWidth = PicWidth / (PicHeight / thumbheight)
        PicHeight = PicHeight / (PicHeight / thumbheight)
    End If
    If ForceFull Then
        If PicWidth < thumbwidth Then
            PicHeight = PicHeight * (thumbwidth / PicWidth)
            PicWidth = thumbwidth
        End If
        If PicHeight < thumbheight Then
            PicWidth = PicWidth * (thumbheight / PicHeight)
            PicHeight = PicHeight * (thumbheight / PicHeight)
        End If
    End If
    temp.cX = PicWidth
    temp.cY = PicHeight
    Thumbsize = temp
End Function

Public Function generatethumbfromimage(picalpha As PictureBox, picbeta As PictureBox, Width As Long, Height As Long, Optional Force As Boolean) As Long
    On Error Resume Next
    Dim PicWidth As Long, PicHeight As Long
    picbeta.Picture = LoadPicture(Empty)
    picbeta.BackColor = TransColor
    
    'picbeta.Move picbeta.Left, picbeta.Top, Width, Height
    
    PicWidth = picalpha.Width
    PicHeight = picalpha.Height
        
    Thumbsize PicWidth, PicHeight, Width, Height, True, Force
    SetStretchBltMode picbeta.hdc, STRETCHMODE
    StretchBlt picbeta.hdc, (picbeta.Width - PicWidth) / 2, (picbeta.Height - PicHeight) / 2, PicWidth, PicHeight, picalpha.hdc, 0, 0, picalpha.Width, picalpha.Height, vbSrcCopy
    
    picbeta.Refresh
    generatethumbfromimage = PicHeight
End Function

Public Function Min(Value1 As Long, Value2 As Long) As Long
    If Value1 < Value2 Then Min = Value1 Else Min = Value2
End Function
Public Function Max(Value1 As Long, Value2 As Long) As Long
    If Value1 < Value2 Then Max = Value2 Else Max = Value1
End Function

Public Function GetValue(Value1 As Long, Value2 As Long, ByVal Percent As Single, Invert As Boolean) As Long
    Dim Lowest As Long, Highest As Long, Difference As Long
    Lowest = Min(Value1, Value2)
    Highest = Max(Value1, Value2)
    Difference = Highest - Lowest
    If Invert Then Percent = 1 - Percent
    GetValue = Difference * Percent + Lowest
End Function










'scroll wheel code
'<0 is ClockWise, >0=counter clockwise
Public Function ScrollMoved(Optional hwnd As Long) As Long
    Dim amsg As msg
    GetMessage amsg, hwnd, 0, 0
    DispatchMessage amsg
    If amsg.message = 522 Then ScrollMoved = amsg.wParam / 65536: DoEvents
End Function
