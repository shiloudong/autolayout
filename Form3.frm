VERSION 5.00
Begin VB.Form ProbeAngleForm 
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
         Caption         =   "Angle"
         Height          =   615
         Left            =   240
         TabIndex        =   14
         Top             =   3120
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
         Left            =   1320
         TabIndex        =   8
         Top             =   3120
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
         Caption         =   "Layer"
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
         TabIndex        =   6
         Top             =   3720
         Width           =   855
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
Attribute VB_Name = "ProbeAngleForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim BL As Double

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
    Call M_RedrawPicutreBox
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
    Call SetPicure(Picture1)
    Picture1.AutoRedraw = True
    M_Index = 0
    Call M_GetExcelData(EntranceForm.TextPath.Text)
    BL = M_GetScale(Picture1.width, Picture1.height)
    F_MovePoint(0) = Picture1.width / 2
    F_MovePoint(1) = Picture1.height / 2
    isMove = False
    isZoom = False
    'move patten
    patten = 2
End Sub

Private Sub MoveCmd_Click()
    patten = 1
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        isMove = True
        previousPoint(0) = x
        previousPoint(1) = y
    ElseIf Button = 2 Then
        isZoom = True
        previousPoint(0) = x
        previousPoint(1) = y
    End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim dX, dY, distance As Double
    If Button = 1 And isMove Then
        If (patten = 1) Then
            dX = x - previousPoint(0)
            dY = y - previousPoint(1)
    
            F_MovePoint(0) = F_MovePoint(0) + dX
            F_MovePoint(1) = F_MovePoint(1) + dY
            Call M_RedrawPicutreBox
            previousPoint(0) = x
            previousPoint(1) = y
        ElseIf patten = 2 Or patten = 3 Then
            Call M_RedrawPicutreBox
            Dim endpoint(0 To 1) As Double
            endpoint(0) = x
            endpoint(1) = y
            Call M_DrawRectangle(previousPoint, endpoint)
        
        End If

    ElseIf Button = 2 And isZoom Then
        
        dX = x - previousPoint(0)
        dY = y - previousPoint(1)
        distance = Sqr(dX * dX + dY * dY)
        If dX > 0 Then
            M_Scale = M_Scale * (1 + distance * 2 / Picture1.width)
        Else
            M_Scale = M_Scale * (1 - distance * 2 / Picture1.width)
        End If
        Call M_RedrawPicutreBox
        previousPoint(0) = x
        previousPoint(1) = y
    End If
    
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim endpoint(0 To 1) As Double
    If Button = 1 Then
        If patten = 1 Then
            isMove = False
        ElseIf patten = 2 Then

            endpoint(0) = x
            endpoint(1) = y
            Call CalculateSelectedPoints(previousPoint, endpoint)
            Call M_RedrawPicutreBox
            Call AngleForm.Show
        ElseIf patten = 3 Then

            endpoint(0) = x
            endpoint(1) = y
            Call CalculateSelectedPoints(previousPoint, endpoint)
            Call M_RedrawPicutreBox
            Call Form5.Show
        End If
    ElseIf Button = 2 Then
        isZoom = False
    End If
End Sub
Private Sub ShowLayerDialog()
    AngleForm.Show
End Sub

Private Sub resetCmd_Click()
    Call M_GetScale(Picture1.width, Picture1.height)
    F_MovePoint(0) = Picture1.width / 2
    F_MovePoint(1) = Picture1.height / 2
    Call M_RedrawPicutreBox
End Sub

Private Sub SelectCmd_Click()
    patten = 2
End Sub

Private Sub UndoCmd_Click()
    patten = 3
End Sub

Private Sub XCmd_Click()
    M_RedrawAngleLine (270)
End Sub
Private Sub ACmd_Click()
    M_RedrawAngleLine (180)
End Sub

Private Sub Picture1_KeyPress(KeyAscii As Integer)
    Call Form_KeyPress(KeyAscii)
End Sub

Private Sub undo()
    If M_Index > 0 Then
        M_Index = M_Index - 1
    End If
    Call M_RedrawPicutreBox
End Sub
Private Sub DCmd_Click()
    M_RedrawAngleLine (360)
End Sub

Private Sub saveCmd_Click()
    Dim corow As Long
    Set excelApp = M_CreateExcel(Form1.TextPath.Text)
    Set excelsheet = excelApp.ActiveWorkbook.Sheets("sheet1")
    For i = 6 To rowCount
        excelsheet.cells(i, 8).value = M_Angles(i - 6)
    Next i
    Call excelApp.Workbooks.Close
End Sub
Private Sub WCmd_Click()
    M_RedrawAngleLine (90)
End Sub
Private Sub QCmd_Click()
    M_RedrawAngleLine (135)
End Sub
Private Sub ECmd_Click()
    M_RedrawAngleLine (45)
End Sub
Private Sub ZCmd_Click()
    M_RedrawAngleLine (225)
End Sub
Private Sub CCmd_Click()
    M_RedrawAngleLine (315)
End Sub


