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
Dim OrderedIndexs() As Integer
Dim selectedCount As Integer
Dim selectedFlag() As Boolean
Dim colorMap(0 To 19) As ColorConstants
Private m_picture As PictureBox

Public Sub SetPicure(pic As PictureBox)
    Set m_picture = pic
End Sub

Private Sub InitializeColorMap()
    colorMap(0) = RGB(255, 255, 255)
    colorMap(1) = RGB(0, 255, 0)
    colorMap(2) = RGB(0, 0, 255)
    colorMap(3) = RGB(255, 0, 255)
    colorMap(4) = RGB(255, 255, 0)
    colorMap(5) = RGB(255, 0, 0)
    colorMap(6) = RGB(0, 255, 255)
    colorMap(7) = RGB(255, 127, 127)
    colorMap(8) = RGB(192, 192, 192)
    colorMap(9) = RGB(127, 255, 191)
    colorMap(10) = RGB(127, 191, 255)
    colorMap(11) = RGB(233, 255, 168)
    colorMap(12) = RGB(255, 255, 255)
    colorMap(13) = RGB(255, 255, 255)
    colorMap(14) = RGB(255, 255, 255)
    colorMap(15) = RGB(255, 255, 255)
    colorMap(16) = RGB(255, 255, 255)
    colorMap(17) = RGB(255, 255, 255)
    colorMap(18) = RGB(255, 255, 255)
    colorMap(19) = RGB(255, 255, 255)
End Sub

Public Sub M_GetExcelData(path As String)
    Set excelApp = M_CreateExcel(EntranceForm.TextPath.Text)
    Set excelsheet = excelApp.ActiveWorkbook.Sheets("sheet1")
    
    Dim PadNo, index As Integer
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
    ReDim Preserve selectedFlag(0 To M_RowCount - 1) As Boolean
    ReDim Preserve OrderedIndexs(0 To M_RowCount - 1) As Integer
    
    Dim layerNo As Integer
    For i = 0 To M_RowCount - 1
        selectedFlag(i) = False
        M_PadNumbers(i) = excelsheet.cells(i + 6, 1).value
        M_PointsX(i) = excelsheet.cells(i + 6, 2).value / 1000
        M_PointsY(i) = excelsheet.cells(i + 6, 3).value / 1000
        M_PadNames(i) = excelsheet.cells(i + 6, 4).value
        M_Traces(i) = excelsheet.cells(i + 6, 5).value
        M_Jumpers(i) = excelsheet.cells(i + 6, 6).value
        M_Channels(i) = excelsheet.cells(i + 6, 7).value
        M_Angles(i) = excelsheet.cells(i + 6, 8).value
        layerNo = excelsheet.cells(i + 6, 9).value
        If (layerNo = 0) Then
            layerNo = 1
        End If
        M_Layers(i) = layerNo
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
    Call InitializeColorMap
End Sub

Public Function M_GetScale(width As Double, height As Double) As Double
    Dim a As Double
    If width < height Then
        a = width
    Else
        a = height
    End If
    If M_MaxX - M_MinX > M_MaxY - M_MinY Then
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

Public Sub M_DrawRectangle(startPoint() As Double, endpoint() As Double)
    m_picture.DrawWidth = 1
    m_picture.Line (startPoint(0), startPoint(1))-(startPoint(0), endpoint(1)), RGB(255, 255, 255)
    m_picture.Line (startPoint(0), startPoint(1))-(endpoint(0), startPoint(1)), RGB(255, 255, 255)
    m_picture.Line (endpoint(0), startPoint(1))-(endpoint(0), endpoint(1)), RGB(255, 255, 255)
    m_picture.Line (startPoint(0), endpoint(1))-(endpoint(0), endpoint(1)), RGB(255, 255, 255)
    
End Sub

Private Sub DrawPoint(x As Double, y As Double, color As ColorConstants)
    m_picture.DrawWidth = 1

    m_picture.Line (x - 1, y - 1)-(x + 1, y + 1), color, B
End Sub
Private Sub DrawAngleLine(point() As Double, index As Integer)
    m_picture.DrawWidth = 1
    Dim pictureAngle As Double
    pictureAngle = -M_Angles(index)
    Dim x1, y1 As Double

    x1 = point(0) + 20 * Cos(3.1415926 * pictureAngle / 180)
    y1 = point(1) + 20 * Sin(3.1415926 * pictureAngle / 180)
    Dim layerIndex As Integer
    layerIndex = M_Layers(index)
    If (layerIndex > 0 And layerIndex < 20) Then
        Dim color As ColorConstants
        color = colorMap(layerIndex - 1)
        m_picture.Line (point(0), point(1))-(x1, y1), color
    End If

End Sub


Private Sub DrawUnit(index As Integer)
    Dim point(0 To 1) As Double
    point(0) = (M_PointsX(index) - M_CenterPoint(0)) * M_Scale + F_MovePoint(0)
    point(1) = (-1 * M_PointsY(index) + M_CenterPoint(1)) * M_Scale + F_MovePoint(1)
    Call DrawAngleLine(point, index)
    
    If selectedFlag(index) Then
        Call DrawPoint(point(0), point(1), RGB(0, 255, 0))
    Else
        Call DrawPoint(point(0), point(1), RGB(255, 255, 255))
    End If
    'Call DrawLayerText(index)
End Sub

Public Sub M_RedrawAngleLine(Angle As Double)
    If M_Index < M_RowCount Then
        M_Angles(M_Index) = Angle
        M_Index = M_Index + 1
        Call M_RedrawPicutreBox
    End If
End Sub
Public Function CalculateSelectedPoints(startPoint() As Double, endpoint() As Double) As Boolean
    Dim inside As Boolean
    Dim existPoints As Boolean
    existPoints = False
    Dim checkPoint(0 To 1) As Double
    For i = 0 To M_RowCount - 1
        checkPoint(0) = (M_PointsX(i) - M_CenterPoint(0)) * M_Scale + F_MovePoint(0)
        checkPoint(1) = (-1 * M_PointsY(i) + M_CenterPoint(1)) * M_Scale + F_MovePoint(1)
        inside = IsInRectange(checkPoint, startPoint, endpoint)
        selectedFlag(i) = inside
        If (inside) Then
            existPoints = True
        End If
        
    Next i
    CalculateSelectedPoints = existPoints
End Function
Private Function IsInRectange(checkPoint() As Double, startPoint() As Double, endpoint() As Double)
    Dim maxX, minX, maxY, minY As Double
    If startPoint(0) > endpoint(0) Then
        maxX = startPoint(0)
        minX = endpoint(0)
    Else
        maxX = endpoint(0)
        minX = startPoint(0)
    End If
    
    If (startPoint(1) > endpoint(1)) Then
        maxY = startPoint(1)
        minY = endpoint(1)
    Else
        maxY = endpoint(1)
        minY = startPoint(1)
    End If
    

    If checkPoint(0) > minX And checkPoint(0) < maxX Then
        If checkPoint(1) > minY And checkPoint(1) < maxY Then
            IsInRectange = True
        Else
            IsInRectange = False
        End If
    Else
        IsInRectange = False
    End If
End Function
Public Sub SetSelectedAngle(Angle As Double)
    For i = 0 To M_RowCount - 1
        If (selectedFlag(i)) Then
            M_Angles(i) = Angle
        End If
    Next i
    Call M_RedrawPicutreBox
End Sub
'order dircetion
'1: X from low to high
'2: X from high to low
'3: Y from low to high
'4: Y from hight to low
Public Sub ReorderLayer(layerArray() As Integer, arrayLength As Integer, orderDirection As Integer)
    Call ReorderSelectedProbe(orderDirection)
    Call UpdateLayers(layerArray, arrayLength)
End Sub

Private Sub ReorderSelectedProbe(direction As Integer)
    Dim index As Integer
    index = 0
    Dim i, j As Integer
    For i = 0 To M_RowCount - 1
        If (selectedFlag(i)) Then
            OrderedIndexs(index) = i
            index = index + 1
        End If
    Next i
    selectedCount = index
    
    Dim tempPoint As Integer
    tempPoint = OrderedIndexs(0)
    Dim isOrder As Boolean
    isOrder = False

    For i = 1 To selectedCount - 1
        For j = 0 To selectedCount - 1 - i
            isOrder = CompareOrder(OrderedIndexs(j), OrderedIndexs(j + 1), direction)
            If isOrder = False Then
                Call SwitchOrder(j, j + 1)
            End If
        Next j

    Next i
    
End Sub
Private Sub SwitchOrder(index1 As Integer, index2 As Integer)
    Dim temp As Integer
    temp = OrderedIndexs(index1)
    OrderedIndexs(index1) = OrderedIndexs(index2)
    OrderedIndexs(index2) = temp
End Sub
'order dircetion
'1: X from low to high
'2: X from high to low
'3: Y from low to high
'4: Y from hight to low
Private Function CompareOrder(index1, index2, direction) As Boolean
    Dim data1, data2 As Double
    Dim isOrder As Boolean
    isOrder = True
    Select Case direction
        Case 1
            data1 = M_PointsX(index1)
            data2 = M_PointsX(index2)
            If (data1 > data2) Then
                isOrder = False
            End If
        Case 2
            data1 = M_PointsX(index1)
            data2 = M_PointsX(index2)
            If (data1 < data2) Then
                isOrder = False
            End If
        Case 3
            data1 = M_PointsY(index1)
            data2 = M_PointsY(index2)
            If (data1 > data2) Then
                isOrder = False
            End If
        Case 4
            data1 = M_PointsY(index1)
            data2 = M_PointsY(index2)
            If (data1 < data2) Then
                isOrder = False
            End If
    End Select
    CompareOrder = isOrder
End Function
Private Sub UpdateLayers(orderArray() As Integer, length As Integer)
    Dim ordersIndex As Integer
    ordersIndex = 0
    For i = 0 To selectedCount - 1
        M_Layers(OrderedIndexs(i)) = orderArray(ordersIndex)
        
        ordersIndex = ordersIndex + 1
        If (ordersIndex = length) Then
            ordersIndex = 0
        End If
    Next i
End Sub



