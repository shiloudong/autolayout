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
Dim pointsX() As Double
Dim pointsY() As Double
Dim angles() As Double
Dim pointsCount As Integer

'Æô¶¯Excel
Private Function CreateExcel(path As String) As Object
    Dim excelApp As Object
    Set excelApp = CreateObject("excel.application")
    excelApp.Workbooks.Open (path)
    Set CreateExcel = excelApp
End Function

Private Sub ClearCmd_Click()
Picture1.Cls
End Sub

Private Sub Form_Activate()
    Call DrawPoints
    pointIndex = 0
End Sub

Private Sub Form_Load()
    Picture1.AutoRedraw = True
End Sub

Private Sub DrawPoints()
    Picture1.DrawWidth = 2
    Dim corow As Long
    Set excelApp = CreateExcel(Form1.TextPath.Text)
    Set excelsheet = excelApp.ActiveWorkbook.Sheets("sheet1")
    corow = excelsheet.usedrange.Rows.count
    'set points array
    pointsCount = corow
    ReDim Preserve pointsX(0 To pointsCount) As Double
    ReDim Preserve pointsY(0 To pointsCount) As Double
    ReDim Preserve angles(0 To pointsCount) As Double
    
    'get the center point of source points
    Dim minX, maxX, minY, maxY As Double
    minX = maxX = excelsheet.cells(6, 2).value
    minY = maxY = excelsheet.cells(6, 3).value
    For i = 7 To corow
        x = excelsheet.cells(i, 2).value
        y = excelsheet.cells(i, 3).value
        If x > maxX Then
            maxX = x
        ElseIf x < minX Then
            minX = x
        End If
        
        If y > maxY Then
            maxY = y
        ElseIf y < minY Then
            minY = y
        End If
    Next i
    
    'source center points
    Dim sX, sY As Double
    Dim c As Double
    Dim a As Double
    Dim b As Double
    a = maxX - minX
    b = maxY - minY
    If a > b Then
    c = a / 1000
    Else
    c = b / 1000
    End If
        
    Dim blscale As Double
    blscale = Picture1.Width * 0.8 / c
    sX = (maxX + minX) / 2000
    sY = (maxY + minY) / 2000
    
    'destination center points
    Dim dX, dY As Double
    dX = Picture1.Width / 2
    dY = Picture1.Height / 2
    
    'draw points in picuture
    For i = 6 To corow
        x = excelsheet.cells(i, 2).value / 1000
        y = excelsheet.cells(i, 3).value / 1000
        Dim x1, y1 As Double
        x1 = (x - sX) * blscale + dX
        y1 = (y - sY) * blscale + dY
        If i = 6 Then
            Picture1.PSet (x1, y1), RGB(255, 0, 0)
        Else
            Picture1.PSet (x1, y1), RGB(0, 255, 0)
        End If
        pointsX(i - 6) = x1
        pointsY(i - 6) = y1
    Next i

    Call excelApp.Workbooks.Close
End Sub
Private Sub DownCmd_Click()
    drawLine (90)
End Sub
Private Sub LeftCmd_Click()
    drawLine (180)
End Sub
Private Sub RightCmd_Click()
    drawLine (0)
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
Private Sub drawLine(angle As Integer)
    Dim x, y, x1, y1 As Double
    x = pointsX(pointIndex)
    y = pointsY(pointIndex)
    x1 = x + 20 * Cos(3.1415926 * angle / 180)
    y1 = y + 20 * Sin(3.1415926 * angle / 180)
    Picture1.Line (x, y)-(x1, y1), RGB(255, 0, 0)
    angles(pointIndex) = angle
    pointIndex = pointIndex + 1
End Sub

