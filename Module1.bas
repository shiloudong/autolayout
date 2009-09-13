Attribute VB_Name = "Module1"

Public M_RowCount As Integer
Public M_PadNumbers() As Integer
Public M_PointsX() As Double
Public M_PointsY() As Double
Public M_PadNames() As String
Public M_Traces() As String
Public M_Jumpers() As String
Public M_Channels() As String
Public M_Angles() As Double
Public M_Layers() As Integer
Public M_CenterPoint(0 To 1) As Double
Public M_MaxX As Double
Public M_MinX As Double
Public M_MaxY As Double
Public M_MinY As Double
Public F_MovePoint(0 To 1) As Double
Public M_Index As Integer
Public M_Scale As Double

Private m_picture As PictureBox


Public Sub SetPicure(pic As PictureBox)
    Set m_picture = pic
End Sub
Public Sub M_GetExcelData(path As String)
    Set excelApp = M_CreateExcel(Form1.TextPath.Text)
    Set excelsheet = excelApp.ActiveWorkbook.Sheets("sheet1")
    
    Dim padNo, index As Integer
    index = 0
    Do While excelsheet.cells(index + 6, 1).value > 0
        index = index + 1
    Loop
    M_RowCount = index
    
    'redefine the length of array
    ReDim Preserve M_PadNumbers(0 To M_RowCount - 1) As Integer
    ReDim Preserve M_PointsX(0 To M_RowCount - 1) As Double
    ReDim Preserve M_PointsY(0 To M_RowCount - 1) As Double
    ReDim Preserve M_PadNames(0 To M_RowCount - 1) As String
    ReDim Preserve M_Traces(0 To M_RowCount - 1) As String
    ReDim Preserve M_Jumpers(0 To M_RowCount - 1) As String
    ReDim Preserve M_Channels(0 To M_RowCount - 1) As String
    ReDim Preserve M_Angles(0 To M_RowCount - 1) As Double
    ReDim Preserve M_Layers(0 To M_RowCount - 1) As Integer
    
    For i = 0 To M_RowCount - 1
        M_PadNumbers(i) = excelsheet.cells(i + 6, 1).value
        M_PointsX(i) = excelsheet.cells(i + 6, 2).value / 1000
        M_PointsY(i) = excelsheet.cells(i + 6, 3).value / 1000
        M_PadNames(i) = excelsheet.cells(i + 6, 4).value
        M_Traces(i) = excelsheet.cells(i + 6, 5).value
        M_Jumpers(i) = excelsheet.cells(i + 6, 6).value
        M_Channels(i) = excelsheet.cells(i + 6, 7).value
        M_Angles(i) = excelsheet.cells(i + 6, 8).value
        M_Layers(i) = excelsheet.cells(i + 6, 9).value
        If i = 0 Then
            M_MaxX = M_PointsX(i)
            M_MinX = M_PointsX(i)
            M_MaxY = M_PointsY(i)
            M_MinY = M_PointsY(i)
        Else
            If M_PointsX(i) > M_MaxX Then
                M_MaxX = M_PointsX(i)
            ElseIf M_PointsX(i) < M_MinX Then
                M_MinX = M_PointsX(i)
            End If
            
            If M_PointsY(i) > M_MaxY Then
                M_MaxY = M_PointsY(i)
            ElseIf M_PointsY(i) < M_MinY Then
                M_MinY = M_PointsY(i)
            End If
        End If
    Next i
    'get center point
    M_CenterPoint(0) = (M_MaxX + M_MinX) / 2
    M_CenterPoint(1) = (M_MaxY + M_MinY) / 2
    '
    Call excelApp.Workbooks.Close
End Sub

Public Function M_GetScale(width As Double, height As Double) As Double
    Dim a As Double
    If width > hight Then
        a = width
    Else
        a = hight
    End If
    If maxX - minX > maxY - minY Then
        M_Scale = a * 0.8 / (M_MaxX - M_MinX)
    Else
        M_Scale = a * 0.8 / (M_MaxY - M_MinY)
    End If
    M_GetScale = M_Scale
End Function

'Æô¶¯Excel
Public Function M_CreateExcel(path As String) As Object
    Dim excelApp As Object
    Set excelApp = CreateObject("excel.application")
    excelApp.Workbooks.Open (path)
    Set M_CreateExcel = excelApp
End Function
Public Sub M_RedrawPicutreBox()
    Call m_picture.Cls
    For i = 0 To M_RowCount - 1
        DrawUnit (i)
    Next i
End Sub

Public Sub M_DrawRectangle(startPoint() As Double, endPoint() As Double)
    m_picture.DrawWidth = 1
    m_picture.Line (startPoint(0), startPoint(1))-(startPoint(0), endPoint(1)), RGB(255, 255, 255)
    m_picture.Line (startPoint(0), startPoint(1))-(endPoint(0), startPoint(1)), RGB(255, 255, 255)
    m_picture.Line (endPoint(0), startPoint(1))-(endPoint(0), endPoint(1)), RGB(255, 255, 255)
    m_picture.Line (startPoint(0), endPoint(1))-(endPoint(0), endPoint(1)), RGB(255, 255, 255)
    
End Sub

Private Sub DrawPoint(X As Double, Y As Double, color As ColorConstants)
    m_picture.DrawWidth = 5
    m_picture.PSet (X, Y), color
End Sub
Private Sub DrawAngleLine(point() As Double, angle As Double)
    m_picture.DrawWidth = 1
    Dim x1, y1 As Double
    If angle <> 0 Then
        x1 = point(0) + 20 * Cos(3.1415926 * angle / 180)
        y1 = point(1) + 20 * Sin(3.1415926 * angle / 180)
        m_picture.Line (point(0), point(1))-(x1, y1), RGB(255, 0, 0)
    End If
End Sub
Private Sub DrawUnit(index As Integer)
    Dim point(0 To 1) As Double
    point(0) = (M_PointsX(index) - M_CenterPoint(0)) * M_Scale + F_MovePoint(0)
    point(1) = (M_PointsY(index) - M_CenterPoint(1)) * M_Scale + F_MovePoint(1)
    Call DrawAngleLine(point, M_Angles(index))
    
    If index <> M_Index Then
        Call DrawPoint(point(0), point(1), RGB(0, 255, 0))
    Else
        Call DrawPoint(point(0), point(1), RGB(255, 0, 0))
    End If
End Sub

Public Sub M_RedrawAngleLine(angle As Double)
    If M_Index < M_RowCount Then
        M_Angles(M_Index) = angle
        M_Index = M_Index + 1
        Call M_RedrawPicutreBox
    End If
End Sub

