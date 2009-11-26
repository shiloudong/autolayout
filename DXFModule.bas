Attribute VB_Name = "Module2"
'启动Autocad函数
Public Function CreateAcad() As IAcadApplication
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
    Set excelApp = M_CreateExcel(ProbeAngleForm.CommonDialog1.FileName)
    Set excelsheet = excelApp.ActiveWorkbook.Sheets("sheet1")
    
    Dim a, b As Double
    Dim BL As Double
    
    'a > b
    a = M_MaxX - M_MinX
    b = M_MaxY - M_MinY
    If a > b Then
    BL = 0.6 * 420 / a
    Else
    BL = 0.6 * 297 / a
    End If
    
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
  
        Dim p(0 To 2) As Double
        Dim x As String
        Dim y As String
        x = M_PointsX(i) * BL
        y = M_PointsY(i) * BL
        p(0) = Val(x) - newCenter(0)
        p(1) = Val(y) - newCenter(1)
        p(2) = 0
        Angle = M_Angles(i)
        tracetext = M_Traces(i)
        padnametext = M_PadNames(i)
        probelayertext = M_Layers(i)
        jumpertext = M_Jumpers(i)
        channeltext = M_Channels(i)
        padnotext = M_PadNumbers(i)
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
    If a > b Then
    Call drawbox(doc, p2, w, l)
    Else
    Call drawbox(doc, p2, l, w)
    End If
'    Call doc.ModelSpace.AddLine(p3, p4)
'    Dim customer As String
'    Dim device As String
'    Dim pins As String
'    customer = excelsheet.cells(1, 2).value
'    device = excelsheet.cells(2, 2).value
'    pins = excelsheet.cells(3, 2).value
'    Call doc.ModelSpace.AddText("Customer:" & customer, p3, 2)
'    Call doc.ModelSpace.AddText("Device:" & device, p5, 2)
'    Call doc.ModelSpace.AddText("Pins:" & pins, p6, 2)
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
   Dim angletext As Double
   pi = 3.1415926
   anglehd = pi * Angle / 180
   Dim probelength As Double
   Dim lettersize As Double
   Dim letteroffset As Double
   lettersize = DXFForm.Text2.Text / 1000
   probelength = DXFForm.Text3.Text / 1000
   letteroffset = DXFForm.Text4.Text / 1000
   
   
   '画矩形和直线
    Dim lay7 As AcadLayer
    Set layer7 = document.Layers.Add("Pads")
    layer7.color = 7
    layer7.Lineweight = 0.3
    document.ActiveLayer = layer7
    Call drawbox(document, centerPoint, 0.04 * BL, 0.04 * BL)
    Dim endpoint(0 To 2) As Double
    endpoint(0) = centerPoint(0) + probelength * BL * Cos(anglehd)
    endpoint(1) = centerPoint(1) + probelength * BL * Sin(anglehd) 'probe的终点坐标xyz
    endpoint(2) = 0
    Call document.ModelSpace.AddLine(centerPoint, endpoint)
    
    'pad No
    Set mytxt = document.TextStyles.Add("mytxt") '添加mytxt样式
    mytxt.fontFile = "c:\windows\fonts\arial.ttf"
    document.ActiveTextStyle = mytxt '将当前文字样式设置为mytxt
    Dim lay1 As AcadLayer
    Set layer1 = document.Layers.Add("PadNo")
    layer1.color = 2
    layer1.Lineweight = 0.3
    document.ActiveLayer = layer1
    Dim padnoposition(0 To 2) As Double
    padnoposition(0) = centerPoint(0) - letteroffset * BL * Cos(anglehd) + lettersize * BL / 2 * Sin(anglehd)
    padnoposition(1) = centerPoint(1) - letteroffset * BL * Sin(anglehd) - lettersize * BL / 2 * Cos(anglehd) 'pad No.的文本放置的位置
    padnoposition(2) = 0
    Set padnoobj = document.ModelSpace.AddText(PadNo, padnoposition, lettersize * BL) '写pad No.的文本
    Call padnoobj.Rotate(padnoposition, anglehd) '旋转pad No.的文本
    
    'pad name
    Dim lay2 As AcadLayer
    Set layer2 = document.Layers.Add("PadName")
    layer2.color = 3
    layer2.Lineweight = 0.3
    document.ActiveLayer = layer2
    Dim padnameposition(0 To 2) As Double
     Dim padnameposition1(0 To 2) As Double
    padnameposition(0) = centerPoint(0) + 0.2 * BL * Cos(anglehd) + lettersize * BL / 2 * Sin(anglehd)
    padnameposition(1) = centerPoint(1) + 0.2 * BL * Sin(anglehd) - lettersize * BL / 2 * Cos(anglehd) 'pad name的文本放置的位置
    padnameposition(2) = 0
    Set padnameobj = document.ModelSpace.AddText(PadName, padnameposition, lettersize * BL) '写pad name的文本
    Call padnameobj.Rotate(padnameposition, anglehd) '旋转pad name的文本
    
    'trace
    Dim lay3 As AcadLayer
    Set layer3 = document.Layers.Add("Trace")
    layer3.color = 4
    layer3.Lineweight = 0.3
    document.ActiveLayer = layer3
    Dim traceposition(0 To 2) As Double
    traceposition(0) = endpoint(0) + 0.02 * BL * Cos(anglehd) + lettersize * BL / 2 * Sin(anglehd)
    traceposition(1) = endpoint(1) + 0.02 * BL * Sin(anglehd) - lettersize * BL / 2 * Cos(anglehd) '焊点的文本放置的位置
    traceposition(2) = 0
    Set traceobj = document.ModelSpace.AddText(Trace, traceposition, lettersize * BL) '写焊点的文本
    Call traceobj.Rotate(traceposition, anglehd) '旋转焊点的文本

    'jumper
    Dim lay4 As AcadLayer
    Set layer4 = document.Layers.Add("Jumper")
    layer4.color = 5
    layer4.Lineweight = 0.3
    document.ActiveLayer = layer4
    Dim jumperposition(0 To 2) As Double
    jumperposition(0) = endpoint(0) + 0.3 * BL * Cos(anglehd) + lettersize * BL / 2 * Sin(anglehd)
    jumperposition(1) = endpoint(1) + 0.3 * BL * Sin(anglehd) - lettersize * BL / 2 * Cos(anglehd) '跳线的文本放置的位置
    jumperposition(2) = 0
    Set jumperobj = document.ModelSpace.AddText(Jumper, jumperposition, lettersize * BL) '写跳线的文本
    Call jumperobj.Rotate(jumperposition, anglehd) '旋转跳线的文本
    
    'channel
    Dim lay5 As AcadLayer
    Set layer5 = document.Layers.Add("Channel")
    layer5.color = 6
    layer5.Lineweight = 0.3
    document.ActiveLayer = layer5
    Dim channelposition(0 To 2) As Double
    channelposition(0) = endpoint(0) + 0.6 * BL * Cos(anglehd) + lettersize * BL / 2 * Sin(anglehd)
    channelposition(1) = endpoint(1) + 0.6 * BL * Sin(anglehd) - lettersize * BL / 2 * Cos(anglehd) 'CH的文本放置的位置
    channelposition(2) = 0
    Set channelobj = document.ModelSpace.AddText(Channel, channelposition, lettersize * BL) '写CH的文本
    Call channelobj.Rotate(channelposition, anglehd) '旋转CH的文本
    
    'layer
    Dim lay6 As AcadLayer
    Set layer6 = document.Layers.Add("Layer")
    layer6.color = 1
    layer6.Lineweight = 0.3
    document.ActiveLayer = layer6
    Dim probelayerposition(0 To 2) As Double
    probelayerposition(0) = centerPoint(0) + 0.1 * BL * Cos(anglehd) + lettersize * BL / 2 * Sin(anglehd)
    probelayerposition(1) = centerPoint(1) + 0.1 * BL * Sin(anglehd) - lettersize * BL / 2 * Cos(anglehd) '层数的文本放置的位置
    probelayerposition(2) = 0
    Set probelayerobj = document.ModelSpace.AddText(Layer, probelayerposition, lettersize * BL) '写层数的文本
    Call probelayerobj.Rotate(probelayerposition, anglehd) '旋转层数的文本
End Sub

'画断面图函数
Public Sub drawsection(document As IAcadDocument, tipdia As Double, tiplength As Double, probedia As Double, taper As Double, theta As Double, beamangle As Double)
Dim p1(0 To 2) As Double
Dim p2(0 To 2) As Double
Dim p3(0 To 2) As Double
Dim p4(0 To 2) As Double
Dim p5(0 To 2) As Double
Dim p6(0 To 2) As Double
Dim p7(0 To 2) As Double
Dim p8(0 To 2) As Double
Dim p9(0 To 2) As Double
Dim p10(0 To 2) As Double
Dim p11(0 To 2) As Double
Dim bendangle As Double
Dim taperhd As Double
Dim thetahd As Double
Dim TL2 As Double
Dim TL1 As Double
Dim EL As Double
Dim EL1 As Double
Dim pi As Double
pi = 3.1415926

taperhd = pi * taper / 360
thetahd = pi * theta / 180
beamanglehd = pi * beamangle / 180
bendanglehd = pi * beamangle / 180
TL = tiplength + 0.002
TL1 = TL / Cos(taperhd)
EL = (probedia - tipdia) / (2 * Tan(taperhd))
EL1 = EL / Cos(taperhd)

p2(0) = 100
p2(1) = 100
p2(2) = 0

p1(0) = p2(0) - tipdia * Cos(thetahd - pi / 2) / 2
p1(1) = p2(1) + tipdia * Sin(thetahd - pi / 2) / 2
p1(2) = 0

p3(0) = p2(0) + tipdia * Cos(thetahd - pi / 2) / 2
p3(1) = p2(1) - tipdia * Sin(thetahd - pi / 2) / 2
p3(2) = 0

p4(0) = p1(0) + TL1 * Cos(pi - thetahd + taperhd)
p4(1) = p1(1) + TL1 * Sin(pi - thetahd + taperhd)
p4(2) = 0

p5(0) = p2(0) + TL * Cos(pi - thetahd)
p5(1) = p2(1) + TL * Sin(pi - thetahd)
p5(2) = 0

p6(0) = p3(0) + TL1 * Cos(pi - thetahd - taperhd)
p6(1) = p3(1) + TL1 * Sin(pi - thetahd - taperhd)
p6(2) = 0

p7(0) = p1(0) + EL1 * Cos(pi - thetahd + taperhd)
p7(1) = p1(1) + EL1 * Sin(pi - thetahd + taperhd)
p7(2) = 0

p8(0) = p2(0) + EL * Cos(pi - thetahd)
p8(1) = p2(1) + EL * Sin(pi - thetahd)
p8(2) = 0

p9(0) = p3(0) + EL1 * Cos(pi - thetahd - taperhd)
p9(1) = p3(1) + EL1 * Sin(pi - thetahd - taperhd)
p9(2) = 0

p10(0) = p7(0) + 60 * Cos(pi - thetahd)
p10(1) = p7(1) + 60 * Sin(pi - thetahd)
p10(2) = 0

p11(0) = p9(0) + 60 * Cos(pi - thetahd)
p11(1) = p9(1) + 60 * Sin(pi - thetahd)
p11(2) = 0

'creat tip
Set line1obj = document.ModelSpace.AddLine(p1, p3)
Set line2obj = document.ModelSpace.AddLine(p1, p4)
Set line3obj = document.ModelSpace.AddLine(p3, p6)
Set line4obj = document.ModelSpace.AddLine(p2, p5)
Set line5obj = document.ModelSpace.AddLine(p4, p6)

'creat 其他
Set line12obj = document.ModelSpace.AddLine(p4, p6)
Set line6obj = document.ModelSpace.AddLine(p4, p7)
Set line7obj = document.ModelSpace.AddLine(p5, p8)
Set line8obj = document.ModelSpace.AddLine(p6, p9)
Set line9obj = document.ModelSpace.AddLine(p7, p9)
Set line10obj = document.ModelSpace.AddLine(p7, p10)
Set line11obj = document.ModelSpace.AddLine(p9, p11)

'Rotate Probe
Call line6obj.Rotate(p6, -(pi - thetahd - beamanglehd))
Call line7obj.Rotate(p6, -(pi - thetahd - beamanglehd))
Call line8obj.Rotate(p6, -(pi - thetahd - beamanglehd))
Call line9obj.Rotate(p6, -(pi - thetahd - beamanglehd))
Call line10obj.Rotate(p6, -(pi - thetahd - beamanglehd))
Call line11obj.Rotate(p6, -(pi - thetahd - beamanglehd))
Call line12obj.Rotate(p6, -(pi - thetahd - beamanglehd))

End Sub

