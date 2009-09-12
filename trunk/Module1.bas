Attribute VB_Name = "Module1"

Public rowCount As Integer
Public padNumbers() As Integer
Public pointsX() As Double
Public pointsY() As Double
Public PadNames() As String
Public Traces() As String
Public Jumpers() As String
Public Channels() As String
Public Angles() As Double
Public Layers() As Integer
Public CenterPoint(0 To 1) As Double
Public maxX As Double
Public minX As Double
Public maxY As Double
Public minY As Double




Public Sub GetExcelData(path As String)
    Set excelApp = CreateExcel(Form1.TextPath.Text)
    Set excelsheet = excelApp.ActiveWorkbook.Sheets("sheet1")
    
    Dim padNo, index As Integer
    index = 0
    Do While excelsheet.cells(index + 6, 1).value > 0
        index = index + 1
    Loop
    rowCount = index
    
    'redefine the length of array
    ReDim Preserve padNumbers(0 To rowCount - 1) As Integer
    ReDim Preserve pointsX(0 To rowCount - 1) As Double
    ReDim Preserve pointsY(0 To rowCount - 1) As Double
    ReDim Preserve PadNames(0 To rowCount - 1) As String
    ReDim Preserve Traces(0 To rowCount - 1) As String
    ReDim Preserve Jumpers(0 To rowCount - 1) As String
    ReDim Preserve Channels(0 To rowCount - 1) As String
    ReDim Preserve Angles(0 To rowCount - 1) As Double
    ReDim Preserve Layers(0 To rowCount - 1) As Integer
   
    For i = 0 To rowCount - 1
        padNumbers(i) = excelsheet.cells(i + 6, 1).value
        pointsX(i) = excelsheet.cells(i + 6, 2).value / 1000
        pointsY(i) = excelsheet.cells(i + 6, 3).value / 1000
        PadNames(i) = excelsheet.cells(i + 6, 4).value
        Traces(i) = excelsheet.cells(i + 6, 5).value
        Jumpers(i) = excelsheet.cells(i + 6, 6).value
        Channels(i) = excelsheet.cells(i + 6, 7).value
        Angles(i) = excelsheet.cells(i + 6, 8).value
        Layers(i) = excelsheet.cells(i + 6, 9).value
        If i = 0 Then
            maxX = pointsX(i)
            minX = pointsX(i)
            maxY = pointsY(i)
            maxY = pointsY(i)
        Else
            If pointsX(i) > maxX Then
                maxX = pointsX(i)
            ElseIf pointsX(i) < minX Then
                minX = pointsX(i)
            End If
            
            If pointsY(i) > maxY Then
                maxY = pointsY(i)
            ElseIf pointsY(i) < minY Then
                minY = pointsY(i)
            End If
        End If
    Next i
    'get center point
    CenterPoint(0) = (maxX + minX) / 2
    CenterPoint(1) = (maxY + minY) / 2
    '
    Call excelApp.Workbooks.Close
End Sub

Public Function GetScale(width As Double, height As Double) As Double
    Dim a As Double
    If width > hight Then
        a = width
    Else
        a = hight
    End If
    If maxX - minX > maxY - minY Then
        GetScale = a * 0.8 / (maxX - minX)
    Else
        GetScale = a * 0.8 / (maxY - minY)
    End If
    
End Function

'Æô¶¯Excel
Private Function CreateExcel(path As String) As Object
    Dim excelApp As Object
    Set excelApp = CreateObject("excel.application")
    excelApp.Workbooks.Open (path)
    Set CreateExcel = excelApp
End Function
