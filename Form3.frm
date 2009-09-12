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
         Left            =   1320
         TabIndex        =   8
         Top             =   8280
         Width           =   975
      End
      Begin VB.CommandButton ClearCmd 
         Caption         =   "Clear"
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
         TabIndex        =   7
         Top             =   8280
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
      Begin VB.CommandButton RightCmd 
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
      Begin VB.CommandButton DownCmd 
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
      Begin VB.CommandButton LeftCmd 
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
Dim tempPoint(0 To 1) As Double
Dim bl As Double



Private Sub ClearCmd_Click()
Picture1.Cls
End Sub

Private Sub Form_Activate()
    Call DrawPoints
    pointIndex = 0
End Sub

Private Sub Form_Load()
    Picture1.AutoRedraw = True
    pointIndex = 0
    Call GetExcelData(Form1.TextPath.Text)
    bl = GetScale(Picture1.width, Picture1.height)
End Sub

Private Sub DrawAngleLines()
    For i = 0 To (pointsCount - 1)
        If pointsX(i) <> 0 And pointsY(i) <> 0 Then
            Call drawPointLine(pointsX(i), pointsY(i), Angles(i))
        End If
    Next i
End Sub

Private Sub DrawPoints()
    Dim point(0 To 1) As Double
    
    'draw points in picuture
    Dim i As Integer
    For i = 0 To rowCount - 1
        Call GetPoint(i)
        Picture1.DrawWidth = 5
        If i = 0 Then
            Picture1.PSet (tempPoint(0), tempPoint(1)), RGB(255, 0, 0)
        Else
            Picture1.PSet (tempPoint(0), tempPoint(1)), RGB(0, 255, 0)
        End If
        Call drawPointLine(tempPoint, Angles(i))
    Next i

End Sub
Private Function GetPoint(index As Integer) As Boolean
    If index < rowCount Then
        tempPoint(0) = (pointsX(index) - CenterPoint(0)) * bl + Picture1.width / 2
        tempPoint(1) = (pointsY(index) - CenterPoint(1)) * bl + Picture1.height / 2
        GetPoint = True
    Else
        GetPoint = False
    End If
End Function

Private Sub DownCmd_Click()
    drawLine (90)
End Sub
Private Sub LeftCmd_Click()
    drawLine (180)
End Sub
Private Sub RightCmd_Click()
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
        Call Redraw(pointIndex)
        Dim point(0 To 1) As Double
        Dim correct As Boolean
        
        'x = pointsX(pointIndex)
        'y = pointsY(pointIndex)
        'x1 = x + 20 * Cos(3.1415926 * angle / 180)
        'y1 = y + 20 * Sin(3.1415926 * angle / 180)
        'Picture1.Line (x, y)-(x1, y1), RGB(255, 0, 0)
        Angles(pointIndex) = angle
        
        correct = GetPoint(pointIndex)
        If correct = True Then
            Call drawPointLine(tempPoint, angle)
        End If
        
        pointIndex = pointIndex + 1
        'go to the first point if went to the end
    
        
        correct = GetPoint(pointIndex)
        If correct = True Then
            Call DrawPoint(tempPoint(0), tempPoint(1), RGB(255, 0, 0))
        End If
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
Private Sub Redraw(index As Integer)
    If index < rowCount Then
        Call Picture1.Cls
        Dim point(0 To 1) As Double
    
        Dim bl As Double
        bl = GetScale(Picture1.width, Picture1.height)
    
        'draw points in picuture
        For i = 0 To rowCount - 1
            point(0) = (pointsX(i) - CenterPoint(0)) * bl + Picture1.width / 2
            point(1) = (pointsY(i) - CenterPoint(1)) * bl + Picture1.height / 2
            If index <> i Then
                Call drawPointLine(point, Angles(i))
            End If
            If index <> i - 1 Then
                Call DrawPoint(point(0), point(1), RGB(0, 255, 0))
            End If
        Next i
    End If

End Sub
Private Sub DrawPoint(x As Double, y As Double, color As ColorConstants)
    Picture1.DrawWidth = 5
    Picture1.PSet (x, y), color
End Sub

