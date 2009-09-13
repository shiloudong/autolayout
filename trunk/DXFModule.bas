Attribute VB_Name = "Module2"
'启动Autocad函数
Private Function CreateAcad() As IAcadApplication
    Dim cad As autoCAD.AcadApplication
    On Error Resume Next
    Set cad = GetObject(, "AutoCAD.Application")
    If Err Then
        Err.Clear
        Set cad = CreateObject("AutoCAD.Application")
        If Err Then End
    End If
    Dim count As Integer
    count = cad.Documents.count
    Set document = cad.ActiveDocument
    cad.Visible = True
    Set CreateAcad = cad
End Function

Public Sub CreateDXFFile()
    Dim app As IAcadApplication
    Dim doc As IAcadDocument
    Set app = CreateAcad
    Set doc = app.Documents.Add

    Set excelApp = M_CreateExcel(EntranceForm.TextPath.Text)
    Set excelsheet = excelApp.ActiveWorkbook.Sheets("sheet1") '当前工作表为sheet1
    
    Dim a, b, c, BL As Double
    a = M_MaxX - M_MinX
    b = M_MaxY - M_MinY
    If a > b Then
        c = a / 1000
    Else
        c = b / 1000
    End If
    BL = 178.2 / c
    
    Dim newCenter(0 To 2) As Double
    newCenter(0) = (M_MaxX + M_MinX) * BL / 2
    newCenter(1) = (M_MaxY + M_MinY) * BL / 2
    newCenter(2) = 0
    Dim Angle As Double
    Dim padnotext As String
    Dim padnametext As String
    Dim tracetext As String
    Dim jumpertext As String
    Dim channeltext As String
    Dim probelayertext As String
    
    For i = 0 To M_RowCount - 1
  
        Dim p(0 To 2) As Double '定义了圆心的位置坐标，下方的p（0），p（1），p（2）为该圆心的x,y,z
        Dim x As String '定义x坐标
        Dim y As String '定义y坐标
        x = M_PointsX(i) * BL '读取excel中的X坐标
        y = M_PointsY(i) * BL '读取excel中的Y坐标
        p(0) = Val(x) - newCenter(0)
        p(1) = Val(y) - newCenter(1)
        p(2) = 0
        Angle = M_Angles(i) '读取excel里的拉针角度
        tracetext = M_Traces(i) '读取excel里的焊点
        padnametext = M_PadNames(i) '读取excel里的pad name
        probelayertext = M_Layers(i) '读取excel里的针层
        jumpertext = M_Jumpers(i) '读取excel里的跳线
        channeltext = M_Channels(i) '读取excel里的CH
        padnotext = M_PadNumbers(i) '读取excel里的Pad No.
        Call DrawUnit(doc, p, Angle, padnotext, padnametext, tracetext, jumpertext, channeltext, probelayertext, BL)
       
    Next i

    'creat layout drawing frame
    Dim lay8 As AcadLayer
    Set layer8 = doc.Layers.Add("Layer")
    layer8.color = 7
    layer8.Lineweight = 0.5
    doc.ActiveLayer = layer8
    Dim p1(0 To 2) As Double
    Dim p2(0 To 2) As Double
    Dim p3(0 To 2) As Double
    Dim p4(0 To 2) As Double
    Dim p5(0 To 2) As Double
    Dim p6(0 To 2) As Double
    Dim l As Double
    Dim w As Double
    l = 297
    w = 420
    p2(0) = 0
    p2(1) = 0
    p2(2) = 0
    Call drawbox(doc, p2, l, w)
    Call doc.ModelSpace.AddLine(p3, p4)
    Dim customer As String
    Dim device As String
    Dim pins As String
    customer = excelsheet.cells(1, 2).value
    device = excelsheet.cells(2, 2).value
    pins = excelsheet.cells(3, 2).value
    'Call doc.ModelSpace.AddText("Customer:" & customer, p3, 250)
    'Call doc.ModelSpace.AddText("Device:" & device, p5, 250)
    'Call doc.ModelSpace.AddText("Pins:" & pins, p6, 250)
    doc.Application.ZoomExtents
    Call excelApp.Workbooks.Close '关闭excel程序
End Sub
'画矩形的函数
Function drawbox(document As IAcadDocument, cp, l, w) '根据矩形中心点坐标画矩形的子程序
    Dim boxp(0 To 14) As Double
    boxp(0) = cp(0) - l / 2
    boxp(1) = cp(1) - w / 2
    boxp(3) = cp(0) - l / 2
    boxp(4) = cp(1) + w / 2
    boxp(6) = cp(0) + l / 2
    boxp(7) = cp(1) + w / 2
    boxp(9) = cp(0) + l / 2
    boxp(10) = cp(1) - w / 2
    boxp(12) = cp(0) - l / 2
    boxp(13) = cp(1) - w / 2
    Call document.ModelSpace.AddPolyline(boxp)
End Function
' 画圆环的函数
Function drawDonut(document As IAcadDocument, D1 As Double, D2 As Double, Pt1 As Variant) As AcadLWPolyline
    Dim LW As Double
    LW = (D1 - D2) / 2
    Dim lwPlineObj As AcadLWPolyline
    Dim points1(0 To 5) As Double
    points1(0) = Pt1(0) - (D1 - LW) / 2
    points1(1) = Pt1(1)
    points1(2) = Pt1(0) + (D1 - LW) / 2
    points1(3) = Pt1(1)
    points1(4) = points1(0)
    points1(5) = points1(1)
    Set lwPlineObj = document.ModelSpace.AddLightWeightPolyline(points1)
    lwPlineObj.SetBulge 0, 1
    lwPlineObj.SetBulge 1, 1
    lwPlineObj.SetWidth 0, LW, LW
    lwPlineObj.SetWidth 1, LW, LW
    Set drawDonut = lwPlineObj
End Function

'画layout的函数
Private Sub DrawUnit(document As IAcadDocument, centerPoint() As Double, Angle As Double, PadNo As String, PadName As String, Trace As String, Jumper As String, Channel As String, Layer As String, BL As Double)
   Dim pi As Double
   Dim anglehd As Double
   pi = 3.1415926
   anglehd = pi * Angle / 180
   
   '画矩形和直线
    Dim lay7 As AcadLayer
    Set layer7 = document.Layers.Add("Pads")
    layer7.color = 7
    layer7.Lineweight = 0.5
    document.ActiveLayer = layer7
    Call drawbox(document, centerPoint, 0.032 * BL, 0.032 * BL)
    Dim endpoint(0 To 2) As Double
    endpoint(0) = centerPoint(0) + EntranceForm.Text3.Text * BL * Cos(anglehd)
    endpoint(1) = centerPoint(1) + EntranceForm.Text3.Text * BL * Sin(anglehd) 'probe的终点坐标xyz
    endpoint(2) = 0
    Call document.ModelSpace.AddLine(centerPoint, endpoint)
    
    'pad No
    Set mytxt = document.TextStyles.Add("mytxt") '添加mytxt样式
    mytxt.fontFile = "c:\windows\fonts\arial.ttf"
    document.ActiveTextStyle = mytxt '将当前文字样式设置为mytxt
    Dim lay1 As AcadLayer
    Set layer1 = document.Layers.Add("PadNo")
    layer1.color = 2
    layer1.Lineweight = 0.5
    document.ActiveLayer = layer1
    Dim padnoposition(0 To 2) As Double
    padnoposition(0) = centerPoint(0) - EntranceForm.Text4.Text * BL * Cos(anglehd) + EntranceForm.Text2.Text * BL / 2 * Sin(anglehd)
    padnoposition(1) = centerPoint(1) - EntranceForm.Text4.Text * BL * Sin(anglehd) - EntranceForm.Text2.Text * BL / 2 * Cos(anglehd) 'pad No.的文本放置的位置
    padnoposition(2) = 0
    Set padnoobj = document.ModelSpace.AddText(PadNo, padnoposition, EntranceForm.Text2.Text * BL) '写pad No.的文本
    Call padnoobj.Rotate(padnoposition, anglehd) '旋转pad No.的文本
    
    'pad name
    Dim lay2 As AcadLayer
    Set layer2 = document.Layers.Add("PadName")
    layer2.color = 3
    layer2.Lineweight = 0.5
    document.ActiveLayer = layer2
    Dim padnameposition(0 To 2) As Double
    padnameposition(0) = centerPoint(0) + 0.2 * BL * Cos(anglehd) + EntranceForm.Text2.Text * BL / 2 * Sin(anglehd)
    padnameposition(1) = centerPoint(1) + 0.2 * BL * Sin(anglehd) - EntranceForm.Text2.Text * BL / 2 * Cos(anglehd) 'pad name的文本放置的位置
    padnameposition(2) = 0
    Set padnameobj = document.ModelSpace.AddText(PadName, padnameposition, EntranceForm.Text2.Text * BL) '写pad name的文本
    Call padnameobj.Rotate(padnameposition, anglehd) '旋转pad name的文本
    
    'trace
    Dim lay3 As AcadLayer
    Set layer3 = document.Layers.Add("Trace")
    layer3.color = 4
    layer3.Lineweight = 0.5
    document.ActiveLayer = layer3
    Dim traceposition(0 To 2) As Double
    traceposition(0) = endpoint(0) + 0.02 * BL * Cos(anglehd) + EntranceForm.Text2.Text * BL / 2 * Sin(anglehd)
    traceposition(1) = endpoint(1) + 0.02 * BL * Sin(anglehd) - EntranceForm.Text2.Text * BL / 2 * Cos(anglehd) '焊点的文本放置的位置
    traceposition(2) = 0
    Set traceobj = document.ModelSpace.AddText(Trace, traceposition, EntranceForm.Text2.Text * BL) '写焊点的文本
    Call traceobj.Rotate(traceposition, anglehd) '旋转焊点的文本

    'jumper
    Dim lay4 As AcadLayer
    Set layer4 = document.Layers.Add("Jumper")
    layer4.color = 5
    layer4.Lineweight = 0.5
    document.ActiveLayer = layer4
    Dim jumperposition(0 To 2) As Double
    jumperposition(0) = endpoint(0) + 0.3 * BL * Cos(anglehd) + EntranceForm.Text2.Text * BL / 2 * Sin(anglehd)
    jumperposition(1) = endpoint(1) + 0.3 * BL * Sin(anglehd) - EntranceForm.Text2.Text * BL / 2 * Cos(anglehd) '跳线的文本放置的位置
    jumperposition(2) = 0
    Set jumperobj = document.ModelSpace.AddText(Jumper, jumperposition, EntranceForm.Text2.Text * BL) '写跳线的文本
    Call jumperobj.Rotate(jumperposition, anglehd) '旋转跳线的文本
    
    'channel
    Dim lay5 As AcadLayer
    Set layer5 = document.Layers.Add("Channel")
    layer5.color = 6
    layer5.Lineweight = 0.5
    document.ActiveLayer = layer5
    Dim channelposition(0 To 2) As Double
    channelposition(0) = endpoint(0) + 0.6 * BL * Cos(anglehd) + EntranceForm.Text2.Text * BL / 2 * Sin(anglehd)
    channelposition(1) = endpoint(1) + 0.6 * BL * Sin(anglehd) - EntranceForm.Text2.Text * BL / 2 * Cos(anglehd) 'CH的文本放置的位置
    channelposition(2) = 0
    Set channelobj = document.ModelSpace.AddText(Channel, channelposition, EntranceForm.Text2.Text * BL) '写CH的文本
    Call channelobj.Rotate(channelposition, anglehd) '旋转CH的文本
    
    'layer
    Dim lay6 As AcadLayer
    Set layer6 = document.Layers.Add("Layer")
    layer6.color = 1
    layer6.Lineweight = 0.5
    document.ActiveLayer = layer6
    Dim probelayerposition(0 To 2) As Double
    probelayerposition(0) = centerPoint(0) + 0.1 * BL * Cos(anglehd) + EntranceForm.Text2.Text * BL / 2 * Sin(anglehd)
    probelayerposition(1) = centerPoint(1) + 0.1 * BL * Sin(anglehd) - EntranceForm.Text2.Text * BL / 2 * Cos(anglehd) '层数的文本放置的位置
    probelayerposition(2) = 0
    Set probelayerobj = document.ModelSpace.AddText(Layer, probelayerposition, EntranceForm.Text2.Text * BL) '写层数的文本
    Call probelayerobj.Rotate(probelayerposition, anglehd) '旋转层数的文本
End Sub
