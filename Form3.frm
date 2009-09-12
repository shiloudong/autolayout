VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Probe Angle"
   ClientHeight    =   9330
   ClientLeft      =   1545
   ClientTop       =   1335
   ClientWidth     =   11970
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   164.571
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   211.138
   Begin VB.Frame Frame1 
      Height          =   9255
      Left            =   9360
      TabIndex        =   1
      Top             =   0
      Width           =   2535
      Begin VB.CommandButton SelectCmd 
         Caption         =   "Select"
         Height          =   615
         Left            =   240
         TabIndex        =   14
         Top             =   3360
         Width           =   855
      End
      Begin VB.CommandButton resetCmd 
         Caption         =   "Reset"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   13
         Top             =   2520
         Width           =   975
      End
      Begin VB.CommandButton Ccmd 
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         TabIndex        =   12
         Top             =   1440
         Width           =   615
      End
      Begin VB.CommandButton Zcmd 
         Caption         =   "Z"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   1440
         Width           =   615
      End
      Begin VB.CommandButton QCmd 
         Caption         =   "Q"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Ecmd 
         Caption         =   "E"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         TabIndex        =   9
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton saveCmd 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1200
         TabIndex        =   8
         Top             =   3360
         Width           =   975
      End
      Begin VB.CommandButton MoveCmd 
         Caption         =   "Move"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1320
         TabIndex        =   7
         Top             =   2520
         Width           =   975
      End
      Begin VB.CommandButton UndoCmd 
         Caption         =   "Undo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         TabIndex        =   6
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton DCmd 
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         TabIndex        =   5
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton XCmd 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         TabIndex        =   4
         Top             =   1440
         Width           =   615
      End
      Begin VB.CommandButton WCmd 
         Caption         =   "W"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         TabIndex        =   3
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton ACmd 
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   615
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   9135
      Left            =   120
      ScaleHeight     =   160.073
      ScaleMode       =   6  'Millimeter
      ScaleWidth      =   160.073
      TabIndex        =   0
      Top             =   120
      Width           =   9135
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim pointIndex As Integer
Dim bl As Double
Dim movePoint(0 To 1) As Double
Dim previousPoint(0 To 1) As Double
Dim startPoint(0 To 1) As Double

Dim isMove As Boolean
Dim isZoom As Boolean
Dim patten As Integer



Private Sub ClearCmd_Click()
Picture1.Cls
End Sub

Private Sub Form_Activate()
    'Call DrawPoints
    pointIndex = 0
    Call RedrawAll
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 113 Then
        Call QCmd_Click
    ElseIf KeyAscii = 119 Then
        Call WCmd_Click
    ElseIf KeyAscii = 101 Then
        Call ECmd_Click
    ElseIf KeyAscii = 97 Then
        Call ACmd_Click
    ElseIf KeyAscii = 100 Then
        Call DCmd_Click
    ElseIf KeyAscii = 122 Then
        Call ZCmd_Click
    ElseIf KeyAscii = 120 Then
        Call XCmd_Click
    ElseIf KeyAscii = 99 Then
        Call CCmd_Click
    End If

End Sub

Private Sub Form_Load()
    Picture1.AutoRedraw = True
    pointIndex = 0
    Call GetExcelData(Form1.TextPath.Text)
    bl = GetScale(Picture1.width, Picture1.height)
    movePoint(0) = Picture1.width / 2
    movePoint(1) = Picture1.height / 2
    isMove = False
    isZoom = False
    'move patten
    patten = 1
End Sub

Private Sub DrawAngleLines()
    For i = 0 To (pointsCount - 1)
        If pointsX(i) <> 0 And pointsY(i) <> 0 Then
            Call drawPointLine(pointsX(i), pointsY(i), Angles(i))
        End If
    Next i
End Sub

Private Sub MoveCmd_Click()
    patten = 1
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        isMove = True
        previousPoint(0) = X
        previousPoint(1) = Y
    ElseIf Button = 2 Then
        isZoom = True
        previousPoint(0) = X
        previousPoint(1) = Y
    End If
End Sub
Private Sub DrawRectangle(startPoint() As Double, endPoint() As Double)
    Picture1.DrawWidth = 1
    Picture1.Line (startPoint(0), startPoint(1))-(startPoint(0), endPoint(1)), RGB(255, 255, 255)
    Picture1.Line (startPoint(0), startPoint(1))-(endPoint(0), startPoint(1)), RGB(255, 255, 255)
    Picture1.Line (endPoint(0), startPoint(1))-(endPoint(0), endPoint(1)), RGB(255, 255, 255)
    Picture1.Line (startPoint(0), endPoint(1))-(endPoint(0), endPoint(1)), RGB(255, 255, 255)
    
End Sub
Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim dX, dY, distance As Double
    If Button = 1 And isMove Then
        If (patten = 1) Then
            dX = X - previousPoint(0)
            dY = Y - previousPoint(1)
    
            movePoint(0) = movePoint(0) + dX
            movePoint(1) = movePoint(1) + dY
            Call RedrawAll
            previousPoint(0) = X
            previousPoint(1) = Y
        ElseIf patten = 2 Then
            Call RedrawAll
            Dim endPoint(0 To 1) As Double
            endPoint(0) = X
            endPoint(1) = Y
            Call DrawRectangle(previousPoint, endPoint)
        End If

    ElseIf Button = 2 And isZoom Then
        
        dX = X - previousPoint(0)
        dY = Y - previousPoint(1)
        distance = Sqr(dX * dX + dY * dY)
        If dX > 0 Then
            bl = bl * (1 + distance * 2 / Picture1.width)
        Else
            bl = bl * (1 - distance * 2 / Picture1.width)
        End If
        Call RedrawAll
        previousPoint(0) = X
        previousPoint(1) = Y
    End If
    
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        isMove = False
    ElseIf Button = 2 Then
        isZoom = False
    End If
End Sub

Private Sub resetCmd_Click()
    bl = GetScale(Picture1.width, Picture1.height)
    movePoint(0) = Picture1.width / 2
    movePoint(1) = Picture1.height / 2
    Call RedrawAll
End Sub

Private Sub SelectCmd_Click()
    patten = 2
End Sub

Private Sub UndoCmd_Click()
    Call undo
End Sub

Private Sub XCmd_Click()
    drawLine (90)
End Sub
Private Sub ACmd_Click()
    drawLine (180)
End Sub

Private Sub Picture1_KeyPress(KeyAscii As Integer)
    Call Form_KeyPress(KeyAscii)
End Sub

Private Sub undo()
    If pointIndex > 0 Then
        pointIndex = pointIndex - 1
    End If
    Call RedrawAll
End Sub
Private Sub DCmd_Click()
    drawLine (360)
End Sub

Private Sub saveCmd_Click()
    Dim corow As Long
    Set excelApp = CreateExcel(Form1.TextPath.Text)
    Set excelsheet = excelApp.ActiveWorkbook.Sheets("sheet1")
    corow = excelsheet.usedrange.Rows.count
    
    For i = 6 To corow
        excelsheet.cells(i, 8).value = Angles(i - 6)
    Next i
    Call excelApp.Workbooks.Close
End Sub
Private Sub WCmd_Click()
    drawLine (270)
End Sub
Private Sub QCmd_Click()
    drawLine (225)
End Sub
Private Sub ECmd_Click()
    drawLine (315)
End Sub
Private Sub ZCmd_Click()
    drawLine (135)
End Sub
Private Sub CCmd_Click()
    drawLine (45)
End Sub

Private Sub drawLine(angle As Double)
    If pointIndex < rowCount Then
        Angles(pointIndex) = angle
        pointIndex = pointIndex + 1
        Call RedrawAll
    End If

End Sub

Private Sub drawPointLine(point() As Double, angle As Double)
    Picture1.DrawWidth = 1
    Dim x1, y1 As Double
    If angle <> 0 Then
        x1 = point(0) + 20 * Cos(3.1415926 * angle / 180)
        y1 = point(1) + 20 * Sin(3.1415926 * angle / 180)
        Picture1.Line (point(0), point(1))-(x1, y1), RGB(255, 0, 0)
    End If

End Sub

Private Sub DrawPoint(X As Double, Y As Double, color As ColorConstants)
    Picture1.DrawWidth = 5
    Picture1.PSet (X, Y), color
End Sub

Private Sub RedrawAll()
    Call Picture1.Cls
    For i = 0 To rowCount - 1
        DrawUnit (i)
    Next i
End Sub

Private Sub DrawUnit(index As Integer)
    Dim point(0 To 1) As Double
    point(0) = (pointsX(index) - CenterPoint(0)) * bl + movePoint(0)
    point(1) = (pointsY(index) - CenterPoint(1)) * bl + movePoint(1)
    Call drawPointLine(point, Angles(index))
    
    If index <> pointIndex Then
        Call DrawPoint(point(0), point(1), RGB(0, 255, 0))
    Else
        Call DrawPoint(point(0), point(1), RGB(255, 0, 0))
    End If
    
End Sub


