Attribute VB_Name = "Module2"
'����Autocad����
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
    Set excelsheet = excelApp.ActiveWorkbook.Sheets("sheet1") '��ǰ������Ϊsheet1
    
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
  
        Dim p(0 To 2) As Double '������Բ�ĵ�λ�����꣬�·���p��0����p��1����p��2��Ϊ��Բ�ĵ�x,y,z
        Dim x As String '����x����
        Dim y As String '����y����
        x = M_PointsX(i) * BL '��ȡexcel�е�X����
        y = M_PointsY(i) * BL '��ȡexcel�е�Y����
        p(0) = Val(x) - newCenter(0)
        p(1) = Val(y) - newCenter(1)
        p(2) = 0
        Angle = M_Angles(i) '��ȡexcel�������Ƕ�
        tracetext = M_Traces(i) '��ȡexcel��ĺ���
        padnametext = M_PadNames(i) '��ȡexcel���pad name
        probelayertext = M_Layers(i) '��ȡexcel������
        jumpertext = M_Jumpers(i) '��ȡexcel�������
        channeltext = M_Channels(i) '��ȡexcel���CH
        padnotext = M_PadNumbers(i) '��ȡexcel���Pad No.
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
    Call excelApp.Workbooks.Close '�ر�excel����
End Sub
'�����εĺ���
Function drawbox(document As IAcadDocument, cp, l, w) '���ݾ������ĵ����껭���ε��ӳ���
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
' ��Բ���ĺ���
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

'��layout�ĺ���
Private Sub DrawUnit(document As IAcadDocument, centerPoint() As Double, Angle As Double, PadNo As String, PadName As String, Trace As String, Jumper As String, Channel As String, Layer As String, BL As Double)
   Dim pi As Double
   Dim anglehd As Double
   pi = 3.1415926
   anglehd = pi * Angle / 180
   
   '�����κ�ֱ��
    Dim lay7 As AcadLayer
    Set layer7 = document.Layers.Add("Pads")
    layer7.color = 7
    layer7.Lineweight = 0.5
    document.ActiveLayer = layer7
    Call drawbox(document, centerPoint, 0.032 * BL, 0.032 * BL)
    Dim endpoint(0 To 2) As Double
    endpoint(0) = centerPoint(0) + EntranceForm.Text3.Text * BL * Cos(anglehd)
    endpoint(1) = centerPoint(1) + EntranceForm.Text3.Text * BL * Sin(anglehd) 'probe���յ�����xyz
    endpoint(2) = 0
    Call document.ModelSpace.AddLine(centerPoint, endpoint)
    
    'pad No
    Set mytxt = document.TextStyles.Add("mytxt") '���mytxt��ʽ
    mytxt.fontFile = "c:\windows\fonts\arial.ttf"
    document.ActiveTextStyle = mytxt '����ǰ������ʽ����Ϊmytxt
    Dim lay1 As AcadLayer
    Set layer1 = document.Layers.Add("PadNo")
    layer1.color = 2
    layer1.Lineweight = 0.5
    document.ActiveLayer = layer1
    Dim padnoposition(0 To 2) As Double
    padnoposition(0) = centerPoint(0) - EntranceForm.Text4.Text * BL * Cos(anglehd) + EntranceForm.Text2.Text * BL / 2 * Sin(anglehd)
    padnoposition(1) = centerPoint(1) - EntranceForm.Text4.Text * BL * Sin(anglehd) - EntranceForm.Text2.Text * BL / 2 * Cos(anglehd) 'pad No.���ı����õ�λ��
    padnoposition(2) = 0
    Set padnoobj = document.ModelSpace.AddText(PadNo, padnoposition, EntranceForm.Text2.Text * BL) 'дpad No.���ı�
    Call padnoobj.Rotate(padnoposition, anglehd) '��תpad No.���ı�
    
    'pad name
    Dim lay2 As AcadLayer
    Set layer2 = document.Layers.Add("PadName")
    layer2.color = 3
    layer2.Lineweight = 0.5
    document.ActiveLayer = layer2
    Dim padnameposition(0 To 2) As Double
    padnameposition(0) = centerPoint(0) + 0.2 * BL * Cos(anglehd) + EntranceForm.Text2.Text * BL / 2 * Sin(anglehd)
    padnameposition(1) = centerPoint(1) + 0.2 * BL * Sin(anglehd) - EntranceForm.Text2.Text * BL / 2 * Cos(anglehd) 'pad name���ı����õ�λ��
    padnameposition(2) = 0
    Set padnameobj = document.ModelSpace.AddText(PadName, padnameposition, EntranceForm.Text2.Text * BL) 'дpad name���ı�
    Call padnameobj.Rotate(padnameposition, anglehd) '��תpad name���ı�
    
    'trace
    Dim lay3 As AcadLayer
    Set layer3 = document.Layers.Add("Trace")
    layer3.color = 4
    layer3.Lineweight = 0.5
    document.ActiveLayer = layer3
    Dim traceposition(0 To 2) As Double
    traceposition(0) = endpoint(0) + 0.02 * BL * Cos(anglehd) + EntranceForm.Text2.Text * BL / 2 * Sin(anglehd)
    traceposition(1) = endpoint(1) + 0.02 * BL * Sin(anglehd) - EntranceForm.Text2.Text * BL / 2 * Cos(anglehd) '������ı����õ�λ��
    traceposition(2) = 0
    Set traceobj = document.ModelSpace.AddText(Trace, traceposition, EntranceForm.Text2.Text * BL) 'д������ı�
    Call traceobj.Rotate(traceposition, anglehd) '��ת������ı�

    'jumper
    Dim lay4 As AcadLayer
    Set layer4 = document.Layers.Add("Jumper")
    layer4.color = 5
    layer4.Lineweight = 0.5
    document.ActiveLayer = layer4
    Dim jumperposition(0 To 2) As Double
    jumperposition(0) = endpoint(0) + 0.3 * BL * Cos(anglehd) + EntranceForm.Text2.Text * BL / 2 * Sin(anglehd)
    jumperposition(1) = endpoint(1) + 0.3 * BL * Sin(anglehd) - EntranceForm.Text2.Text * BL / 2 * Cos(anglehd) '���ߵ��ı����õ�λ��
    jumperposition(2) = 0
    Set jumperobj = document.ModelSpace.AddText(Jumper, jumperposition, EntranceForm.Text2.Text * BL) 'д���ߵ��ı�
    Call jumperobj.Rotate(jumperposition, anglehd) '��ת���ߵ��ı�
    
    'channel
    Dim lay5 As AcadLayer
    Set layer5 = document.Layers.Add("Channel")
    layer5.color = 6
    layer5.Lineweight = 0.5
    document.ActiveLayer = layer5
    Dim channelposition(0 To 2) As Double
    channelposition(0) = endpoint(0) + 0.6 * BL * Cos(anglehd) + EntranceForm.Text2.Text * BL / 2 * Sin(anglehd)
    channelposition(1) = endpoint(1) + 0.6 * BL * Sin(anglehd) - EntranceForm.Text2.Text * BL / 2 * Cos(anglehd) 'CH���ı����õ�λ��
    channelposition(2) = 0
    Set channelobj = document.ModelSpace.AddText(Channel, channelposition, EntranceForm.Text2.Text * BL) 'дCH���ı�
    Call channelobj.Rotate(channelposition, anglehd) '��תCH���ı�
    
    'layer
    Dim lay6 As AcadLayer
    Set layer6 = document.Layers.Add("Layer")
    layer6.color = 1
    layer6.Lineweight = 0.5
    document.ActiveLayer = layer6
    Dim probelayerposition(0 To 2) As Double
    probelayerposition(0) = centerPoint(0) + 0.1 * BL * Cos(anglehd) + EntranceForm.Text2.Text * BL / 2 * Sin(anglehd)
    probelayerposition(1) = centerPoint(1) + 0.1 * BL * Sin(anglehd) - EntranceForm.Text2.Text * BL / 2 * Cos(anglehd) '�������ı����õ�λ��
    probelayerposition(2) = 0
    Set probelayerobj = document.ModelSpace.AddText(Layer, probelayerposition, EntranceForm.Text2.Text * BL) 'д�������ı�
    Call probelayerobj.Rotate(probelayerposition, anglehd) '��ת�������ı�
End Sub
