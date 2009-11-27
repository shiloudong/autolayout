VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form ProbeAngleForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AutoLayout"
   ClientHeight    =   9030
   ClientLeft      =   2715
   ClientTop       =   1620
   ClientWidth     =   13935
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   159.279
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   245.798
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   12120
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   7215
      Begin VB.CommandButton SectionCmd 
         Caption         =   "Section"
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
         Left            =   5400
         TabIndex        =   9
         Top             =   160
         Width           =   855
      End
      Begin VB.CommandButton MaskCmd 
         Caption         =   "MASK"
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
         Left            =   4560
         TabIndex        =   8
         Top             =   160
         Width           =   735
      End
      Begin VB.CommandButton DXFCmd 
         Caption         =   "DXF"
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
         Left            =   3720
         TabIndex        =   7
         Top             =   160
         Width           =   735
      End
      Begin VB.CommandButton SelectCmd 
         Caption         =   "Angle"
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
         Left            =   2040
         TabIndex        =   6
         Top             =   160
         Width           =   735
      End
      Begin VB.CommandButton resetCmd 
         Caption         =   "Zoom All"
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
         TabIndex        =   5
         Top             =   160
         Width           =   975
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
         Height          =   495
         Left            =   6360
         TabIndex        =   4
         Top             =   160
         Width           =   735
      End
      Begin VB.CommandButton ZoomCmd 
         Caption         =   "Zoom"
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
         Left            =   120
         TabIndex        =   3
         Top             =   160
         Width           =   735
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
         Height          =   495
         Left            =   2880
         TabIndex        =   2
         Top             =   160
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   7935
      Left            =   120
      ScaleHeight     =   138.906
      ScaleMode       =   6  'Millimeter
      ScaleWidth      =   240.506
      TabIndex        =   0
      Top             =   960
      Width           =   13695
   End
   Begin VB.Menu File 
      Caption         =   "File(&F)"
      Index           =   1
      Begin VB.Menu LoadExcel 
         Caption         =   "Load Excel"
         Shortcut        =   ^E
      End
      Begin VB.Menu ExportMask 
         Caption         =   "Export Mask"
         Shortcut        =   ^A
      End
      Begin VB.Menu ExportLayout 
         Caption         =   "Export Layout"
         Shortcut        =   ^C
      End
      Begin VB.Menu SectionDrawing 
         Caption         =   "Secton Drawing"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu About 
      Caption         =   "About(&A)"
      Begin VB.Menu AboutAutolayout 
         Caption         =   "About AutoLayout"
      End
   End
End
Attribute VB_Name = "ProbeAngleForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gActiveDoc As IAcadDocument
Dim gAcadApplication As IAcadApplication
Dim BL As Double
Dim previousPoint(0 To 1) As Double
Dim startPoint(0 To 1) As Double
Dim isMove As Boolean
Dim isZoom As Boolean
Dim patten As Integer

Private Sub AboutAutolayout_Click()
AboutForm.Show
End Sub

Private Sub DXFCmd_Click()
    DXFForm.Show
End Sub

Private Sub ExportLayout_Click()
DXFForm.Show
End Sub

Private Sub ExportMask_Click()
MaskForm.Show
End Sub

'Private Sub Form_Activate()
'    Call DrawPoints
'    pointIndex = 0
'    Call M_RedrawPicutreBox
'End Sub

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

'Private Sub Form_Load()
'    Call SetPicure(Picture1)
'    Picture1.AutoRedraw = True
'    M_Index = 0
'    Call M_GetExcelData(ProbeAngleForm.CommonDialog1.FileName)
'    BL = M_GetScale(Picture1.width, Picture1.height)
'    F_MovePoint(0) = Picture1.width / 2
'    F_MovePoint(1) = Picture1.height / 2
'    isMove = False
'    isZoom = False
'    'move patten
'    patten = 2
'End Sub

Private Sub LoadExcel_Click()
   On Error GoTo ErrHandler
    CommonDialog1.Filter = "excelfile (*.xls)|*.xls|"
    CommonDialog1.ShowOpen
    'TextPath.Text = CommonDialog1.FileName
    Call SetPicure(Picture1)
    Picture1.AutoRedraw = True
    M_Index = 0
    Call M_GetExcelData(ProbeAngleForm.CommonDialog1.FileName)
    BL = M_GetScale(Picture1.width, Picture1.height)
    F_MovePoint(0) = Picture1.width / 2
    F_MovePoint(1) = Picture1.height / 2
    isMove = False
    isZoom = False
    'move patten
    patten = 2
    Call resetCmd_Click
    Exit Sub
ErrHandler:
    Exit Sub
End Sub

Private Sub MaskCmd_Click()
MaskForm.Show
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        isZoom = True
        previousPoint(0) = x
        previousPoint(1) = y
    ElseIf Button = 2 Then
        isMove = True
        previousPoint(0) = x
        previousPoint(1) = y
    End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim dX, dY, distance As Double
    If Button = 1 And isZoom Then
        If (patten = 1) Then
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
        ElseIf patten = 2 Or patten = 3 Then
            Call M_RedrawPicutreBox
            Dim endpoint(0 To 1) As Double
            endpoint(0) = x
            endpoint(1) = y
            Call M_DrawRectangle(previousPoint, endpoint)
        End If

    ElseIf Button = 2 And isMove Then
        
        dX = x - previousPoint(0)
        dY = y - previousPoint(1)
        F_MovePoint(0) = F_MovePoint(0) + dX
        F_MovePoint(1) = F_MovePoint(1) + dY
        Call M_RedrawPicutreBox
        previousPoint(0) = x
        previousPoint(1) = y
    End If
    
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim endpoint(0 To 1) As Double
    Dim exist As Boolean
    
    If Button = 1 Then
        If patten = 1 Then
            isZoom = False
            
        ElseIf patten = 2 Then

            endpoint(0) = x
            endpoint(1) = y
            exist = CalculateSelectedPoints(previousPoint, endpoint)
            Call M_RedrawPicutreBox
            If (exist) Then
                Call AngleForm.Show
            End If

        ElseIf patten = 3 Then

            endpoint(0) = x
            endpoint(1) = y
            exist = CalculateSelectedPoints(previousPoint, endpoint)
            Call M_RedrawPicutreBox
            If (exist) Then
                Call LayerForm.Show
            End If

        End If
    ElseIf Button = 2 Then
        isMove = False
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

Private Sub SectionCmd_Click()
SectionForm.Show
End Sub

Private Sub SectionDrawing_Click()
SectionForm.Show
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
    
    Set excelApp = M_CreateExcel(ProbeAngleForm.CommonDialog1.FileName)
    Set excelsheet = excelApp.ActiveWorkbook.Sheets("sheet1")
    For i = 0 To M_RowCount - 1
        excelsheet.cells(i + 6, 8).value = M_Angles(i)
        excelsheet.cells(i + 6, 9).value = M_Layers(i)
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


Private Sub ZoomCmd_Click()
    patten = 1
End Sub
