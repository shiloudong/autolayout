VERSION 5.00
Begin VB.Form EntranceForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Probe Card Design - AutoLayout"
   ClientHeight    =   6225
   ClientLeft      =   3540
   ClientTop       =   1935
   ClientWidth     =   6345
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   6345
   Begin VB.CommandButton Command2 
      Caption         =   "About"
      Height          =   255
      Left            =   5400
      TabIndex        =   22
      Top             =   70
      Width           =   855
   End
   Begin VB.Frame Frame3 
      Caption         =   "Needle Assembly"
      Height          =   1215
      Left            =   120
      TabIndex        =   16
      Top             =   4920
      Width           =   6120
      Begin VB.TextBox Text14 
         Height          =   285
         Left            =   1440
         TabIndex        =   25
         Text            =   "94"
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Section 
         Caption         =   "Creat Section Drawing"
         Height          =   375
         Left            =   3360
         TabIndex        =   17
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Browse Excel of Needle Force First "
         Height          =   255
         Left            =   2760
         TabIndex        =   27
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label14 
         Caption         =   "Theta"
         Height          =   255
         Left            =   600
         TabIndex        =   26
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "MASK"
      Height          =   2055
      Left            =   3240
      TabIndex        =   6
      Top             =   2760
      Width           =   3000
      Begin VB.TextBox Text5 
         Height          =   270
         Left            =   1560
         TabIndex        =   24
         Text            =   "25"
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton mask 
         Caption         =   "Creat Mask"
         Height          =   375
         Left            =   1560
         TabIndex        =   12
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox Text11 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1560
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox Text10 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1560
         TabIndex        =   9
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Width [um]"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Offset [um]"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Diameter [um]"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "LAYOUT"
      Height          =   2055
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Width           =   3000
      Begin VB.TextBox Text4 
         Height          =   270
         Left            =   1440
         TabIndex        =   21
         Text            =   "0.1"
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   270
         Left            =   1440
         TabIndex        =   19
         Text            =   "0.7"
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   1440
         TabIndex        =   15
         Text            =   "0.03"
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton fan 
         Caption         =   "Probe Angle"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton layout 
         Caption         =   "Creat Layout"
         Height          =   375
         Left            =   1560
         TabIndex        =   11
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Offset [um]"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Length [um]"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Font 
         Caption         =   "Font [um]"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.TextBox TextPath 
      Height          =   300
      Left            =   3240
      TabIndex        =   3
      Top             =   480
      Width           =   3000
   End
   Begin VB.DirListBox Dir1 
      Height          =   1620
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   3000
   End
   Begin VB.DriveListBox Drive1 
      Height          =   345
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3000
   End
   Begin VB.FileListBox File1 
      Height          =   1665
      Left            =   3240
      Pattern         =   "*.xls"
      TabIndex        =   0
      Top             =   960
      Width           =   3000
   End
   Begin VB.Label Label3 
      Caption         =   "Browse the Excel File"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "EntranceForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gActiveDoc As IAcadDocument
Dim gAcadApplication As IAcadApplication
Private Sub Command1_Click()
    File1.Pattern = "*.xls"
End Sub
Private Sub Command2_Click()
    Dim newform As New Form2
    newform.Show
End Sub
Private Sub Dir1_Change()
    File1.path = Dir1.path
End Sub
Private Sub Drive1_Change()
    Dir1.path = Drive1.Drive
End Sub
Private Sub fan_Click()
    Dim newform As New Form3
    newform.Show
End Sub
Private Sub File1_Click()
    TextPath.Text = File1.path + "\" + File1.FileName
End Sub
Private Sub Form_Load()
    TextPath.Text = "pa.xls"
End Sub

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

'启动Excel函数
Private Function CreateExcel(path As String) As Object
    Dim excelApp As Object
    Set excelApp = CreateObject("excel.application")
    excelApp.Workbooks.Open (path)
    Set CreateExcel = excelApp
End Function

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
    endpoint(0) = centerPoint(0) + Text3.Text * BL * Cos(anglehd)
    endpoint(1) = centerPoint(1) + Text3.Text * BL * Sin(anglehd) 'probe的终点坐标xyz
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
    padnoposition(0) = centerPoint(0) - Text4.Text * BL * Cos(anglehd) + Text2.Text * BL / 2 * Sin(anglehd)
    padnoposition(1) = centerPoint(1) - Text4.Text * BL * Sin(anglehd) - Text2.Text * BL / 2 * Cos(anglehd) 'pad No.的文本放置的位置
    padnoposition(2) = 0
    Set padnoobj = document.ModelSpace.AddText(PadNo, padnoposition, Text2.Text * BL) '写pad No.的文本
    Call padnoobj.Rotate(padnoposition, anglehd) '旋转pad No.的文本
    
    'pad name
    Dim lay2 As AcadLayer
    Set layer2 = document.Layers.Add("PadName")
    layer2.color = 3
    layer2.Lineweight = 0.5
    document.ActiveLayer = layer2
    Dim padnameposition(0 To 2) As Double
    padnameposition(0) = centerPoint(0) + 0.2 * BL * Cos(anglehd) + Text2.Text * BL / 2 * Sin(anglehd)
    padnameposition(1) = centerPoint(1) + 0.2 * BL * Sin(anglehd) - Text2.Text * BL / 2 * Cos(anglehd) 'pad name的文本放置的位置
    padnameposition(2) = 0
    Set padnameobj = document.ModelSpace.AddText(PadName, padnameposition, Text2.Text * BL) '写pad name的文本
    Call padnameobj.Rotate(padnameposition, anglehd) '旋转pad name的文本
    
    'trace
    Dim lay3 As AcadLayer
    Set layer3 = document.Layers.Add("Trace")
    layer3.color = 4
    layer3.Lineweight = 0.5
    document.ActiveLayer = layer3
    Dim traceposition(0 To 2) As Double
    traceposition(0) = endpoint(0) + 0.02 * BL * Cos(anglehd) + Text2.Text * BL / 2 * Sin(anglehd)
    traceposition(1) = endpoint(1) + 0.02 * BL * Sin(anglehd) - Text2.Text * BL / 2 * Cos(anglehd) '焊点的文本放置的位置
    traceposition(2) = 0
    Set traceobj = document.ModelSpace.AddText(Trace, traceposition, Text2.Text * BL) '写焊点的文本
    Call traceobj.Rotate(traceposition, anglehd) '旋转焊点的文本

    'jumper
    Dim lay4 As AcadLayer
    Set layer4 = document.Layers.Add("Jumper")
    layer4.color = 5
    layer4.Lineweight = 0.5
    document.ActiveLayer = layer4
    Dim jumperposition(0 To 2) As Double
    jumperposition(0) = endpoint(0) + 0.3 * BL * Cos(anglehd) + Text2.Text * BL / 2 * Sin(anglehd)
    jumperposition(1) = endpoint(1) + 0.3 * BL * Sin(anglehd) - Text2.Text * BL / 2 * Cos(anglehd) '跳线的文本放置的位置
    jumperposition(2) = 0
    Set jumperobj = document.ModelSpace.AddText(Jumper, jumperposition, Text2.Text * BL) '写跳线的文本
    Call jumperobj.Rotate(jumperposition, anglehd) '旋转跳线的文本
    
    'channel
    Dim lay5 As AcadLayer
    Set layer5 = document.Layers.Add("Channel")
    layer5.color = 6
    layer5.Lineweight = 0.5
    document.ActiveLayer = layer5
    Dim channelposition(0 To 2) As Double
    channelposition(0) = endpoint(0) + 0.6 * BL * Cos(anglehd) + Text2.Text * BL / 2 * Sin(anglehd)
    channelposition(1) = endpoint(1) + 0.6 * BL * Sin(anglehd) - Text2.Text * BL / 2 * Cos(anglehd) 'CH的文本放置的位置
    channelposition(2) = 0
    Set channelobj = document.ModelSpace.AddText(Channel, channelposition, Text2.Text * BL) '写CH的文本
    Call channelobj.Rotate(channelposition, anglehd) '旋转CH的文本
    
    'layer
    Dim lay6 As AcadLayer
    Set layer6 = document.Layers.Add("Layer")
    layer6.color = 1
    layer6.Lineweight = 0.5
    document.ActiveLayer = layer6
    Dim probelayerposition(0 To 2) As Double
    probelayerposition(0) = centerPoint(0) + 0.1 * BL * Cos(anglehd) + Text2.Text * BL / 2 * Sin(anglehd)
    probelayerposition(1) = centerPoint(1) + 0.1 * BL * Sin(anglehd) - Text2.Text * BL / 2 * Cos(anglehd) '层数的文本放置的位置
    probelayerposition(2) = 0
    Set probelayerobj = document.ModelSpace.AddText(Layer, probelayerposition, Text2.Text * BL) '写层数的文本
    Call probelayerobj.Rotate(probelayerposition, anglehd) '旋转层数的文本
End Sub

'画断面图函数
Private Sub drawsection(document As IAcadDocument, tipdia As Double, tiplength As Double, probedia As Double, taper As Double, theta As Double, beamangle As Double)
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

'Creat Layout
Private Sub layout_Click()
    Dim app As IAcadApplication
    Dim doc As IAcadDocument
    Set app = CreateAcad
    Set doc = app.Documents.Add
    'If doc.Active = False Then End
    
    Dim corow As Long
    Set excelApp = CreateExcel(TextPath.Text)
    Set excelsheet = excelApp.ActiveWorkbook.Sheets("sheet1") '当前工作表为sheet1
    corow = excelsheet.usedrange.Rows.count '计算工作表的总行数
    Dim maxX, minX, maxY, minY As Double
    minX = excelsheet.cells(6, 2).value  '读取excel中的X坐标
    minY = excelsheet.cells(6, 3).value  '读取excel中的Y坐标
    maxX = excelsheet.cells(6, 2).value  '读取excel中的X坐标
    maxY = excelsheet.cells(6, 3).value  '读取excel中的Y坐标

For i = 7 To corow
    Dim currentX, currentY As Double
    currentX = excelsheet.cells(i, 2).value
    currentY = excelsheet.cells(i, 3).value
    If (currentX < minX) Then
       minX = currentX
    Else
        If (currentX > maxX) Then
            maxX = currentX
        End If
    End If
    
    If (currentY < minY) Then
        minY = currentY
    Else
        If (currentY > maxY) Then
            maxY = currentY
        End If
    End If
Next i

Dim a, b, c, BL As Double
a = maxX - minX
b = maxY - minY
If a > b Then
c = a / 1000
Else
c = b / 1000
End If
BL = 178.2 / c

    Dim newCenter(0 To 2) As Double
    newCenter(0) = (maxX + minX) * BL / 2000
    newCenter(1) = (maxY + minY) * BL / 2000
    newCenter(2) = 0
For i = 6 To corow '循环开始
    Dim Angle As Double '定义拉针角度为double
    Dim tracetext As String
    Dim padnametext As String
    Dim probelayertext As String
    Dim jumpertext As String
    Dim channeltext As String
    Dim padnotext As String
    Dim p(0 To 2) As Double '定义了圆心的位置坐标，下方的p（0），p（1），p（2）为该圆心的x,y,z
    Dim x As String '定义x坐标
    Dim y As String '定义y坐标
    x = excelsheet.cells(i, 2).value * BL / 1000 '读取excel中的X坐标
    y = excelsheet.cells(i, 3).value * BL / 1000 '读取excel中的Y坐标
    p(0) = Val(x) - newCenter(0)
    p(1) = Val(y) - newCenter(1)
    p(2) = 0
    Angle = excelsheet.cells(i, 8).value '读取excel里的拉针角度
    tracetext = excelsheet.cells(i, 5).value '读取excel里的焊点
    padnametext = excelsheet.cells(i, 4).value '读取excel里的pad name
    probelayertext = excelsheet.cells(i, 9).value '读取excel里的针层
    jumpertext = excelsheet.cells(i, 6).value '读取excel里的跳线
    channeltext = excelsheet.cells(i, 7).value '读取excel里的CH
    padnotext = excelsheet.cells(i, 1).value '读取excel里的Pad No.
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

'Creat Mask
Private Sub mask_Click()
    Dim app As IAcadApplication
    Dim doc2 As IAcadDocument
    Set app = CreateAcad
    Set doc2 = app.Documents.Add
    
    Dim maskp(0 To 2) As Double
    Dim ed As Double
    Dim id As Double

    id = Text10.Text / 1000
    ed = Text10.Text / 1000 + Text5.Text / 1000
    Dim corow As Long
    Set excelApp = CreateExcel(TextPath.Text)
    Set excelsheet = excelApp.ActiveWorkbook.Sheets("sheet1") '当前工作表为sheet1
    corow = excelsheet.usedrange.Rows.count '计算工作表的总行数
    Dim Angle As Double

    
    Dim maxX, minX, maxY, minY As Double
    minX = excelsheet.cells(6, 2).value  '读取excel中的X坐标
    minY = excelsheet.cells(6, 3).value  '读取excel中的Y坐标
    maxX = excelsheet.cells(6, 2).value  '读取excel中的X坐标
    maxY = excelsheet.cells(6, 3).value  '读取excel中的Y坐标

For i = 7 To corow
    Dim currentX, currentY As Double
    currentX = excelsheet.cells(i, 2).value
    currentY = excelsheet.cells(i, 3).value
    If (currentX < minX) Then
       minX = currentX
    Else
        If (currentX > maxX) Then
            maxX = currentX
        End If
    End If
    
    If (currentY < minY) Then
        minY = currentY
    Else
        If (currentY > maxY) Then
            maxY = currentY
        End If
        
    End If
Next i

    Dim newCenter(0 To 2) As Double
    newCenter(0) = (maxX + minX) / 2
    newCenter(1) = (maxY + minY) / 2
    newCenter(2) = 0
For i = 6 To corow
    Dim x As String '定义x坐标
    Dim y As String '定义y坐标
    x = excelsheet.cells(i, 2).value '读取excel中的X坐标
    y = excelsheet.cells(i, 3).value '读取excel中的Y坐标
    Angle = excelsheet.cells(i, 8).value '读取excel里的拉针角度
        
    maskp(0) = (Val(x) - newCenter(0)) / 1000 + Text11.Text / 1000 * Cos(3.1415926 * Angle / 180)
    maskp(1) = (Val(y) - newCenter(1)) / 1000 + Text11.Text / 1000 * Sin(3.1415926 * Angle / 180)
    maskp(2) = 0
    Call drawDonut(doc2, ed, id, maskp)
Next i

'Creat mask drawing frame
    Dim framecenter(0 To 2) As Double
    framecenter(0) = 0
    framecenter(1) = -5
    framecenter(2) = 0
    Call drawbox(doc2, framecenter, 30, 30)
    Call drawbox(doc2, framecenter, 50, 50)
    Dim p9(0 To 2) As Double
    Dim p10(0 To 2) As Double
    Dim p11(0 To 2) As Double
    Dim p12(0 To 2) As Double
    Dim p13(0 To 2) As Double
    p9(0) = -10
    p9(1) = -10
    p9(2) = 0
    p10(0) = -10
    p10(1) = -12
    p10(2) = 0
    p11(0) = -10
    p11(1) = -14
    p11(2) = 0
    p12(0) = -10
    p12(1) = -16
    p12(2) = 0
    p13(0) = 0
    p13(1) = -16
    p13(2) = 0
    Dim customer As String
    Dim device As String
    Dim pins As String
    customer = excelsheet.cells(1, 2).value
    device = excelsheet.cells(2, 2).value
    pins = excelsheet.cells(3, 2).value
    Set mytxt = doc2.TextStyles.Add("mytxt") '添加mytxt样式
    mytxt.fontFile = "c:\windows\fonts\arial.ttf"
    doc2.ActiveTextStyle = mytxt '将当前文字样式设置为mytxt
    Call doc2.ModelSpace.AddText("Customer:" & customer, p9, 1.5)
    Call doc2.ModelSpace.AddText("Device:" & device, p10, 1.5)
    Call doc2.ModelSpace.AddText("Pins:" & pins, p11, 1.5)
    Call doc2.ModelSpace.AddText("Dia=" & Text10.Text, p12, 1.5)
    Call doc2.ModelSpace.AddText("Offset=" & Text11.Text, p13, 1.5)
    doc2.Application.ZoomExtents
    Call excelApp.Workbooks.Close '关闭excel程序
End Sub

Private Sub Section_Click()
Dim app As IAcadApplication
    Dim doc3 As IAcadDocument
    Set app = CreateAcad
    Set doc3 = app.Documents.Add
    Dim taper As Double
    Dim beamangle As Double
    Dim bendangle As Double
    Dim theta As Double
    Dim TL2 As Double
    Dim TD As Double
    Dim PD As Double
    theta = Text14.Text
    Set excelApp = CreateExcel(TextPath.Text)
    Set excelsheet = excelApp.ActiveWorkbook.Sheets("sheet1")
For i = 21 To 32
    Dim value As Double
    value = excelsheet.cells(i, 7).value
    If value <> 0 Then
        TD = value / 1000
        TL2 = excelsheet.cells(i, 6).value / 1000
        PD = excelsheet.cells(i, 3).value
        taper = excelsheet.cells(i, 8).value
        beamangle = excelsheet.cells(i, 4).value
        Call drawsection(doc3, TD, TL2, PD, taper, theta, beamangle)
    End If
Next i
doc3.Application.ZoomExtents
Call excelApp.Workbooks.Close
End Sub
