VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form EntranceForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Probe Card Design - AutoLayout"
   ClientHeight    =   4065
   ClientLeft      =   2700
   ClientTop       =   1650
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
   ScaleHeight     =   4065
   ScaleWidth      =   6345
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5760
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FilterIndex     =   2
   End
   Begin VB.Frame Frame3 
      Caption         =   "Needle Assembly"
      Height          =   1215
      Left            =   120
      TabIndex        =   11
      Top             =   2760
      Width           =   6120
      Begin VB.TextBox Text14 
         Height          =   285
         Left            =   1440
         TabIndex        =   19
         Text            =   "94"
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Section 
         Caption         =   "SECTION DRAWING"
         Height          =   375
         Left            =   3360
         TabIndex        =   12
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Browse Excel of Needle Force First "
         Height          =   255
         Left            =   2760
         TabIndex        =   21
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label14 
         Caption         =   "Theta"
         Height          =   255
         Left            =   600
         TabIndex        =   20
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "MASK"
      Height          =   2055
      Left            =   3240
      TabIndex        =   2
      Top             =   600
      Width           =   3000
      Begin VB.TextBox Text5 
         Height          =   270
         Left            =   1560
         TabIndex        =   18
         Text            =   "25"
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton mask 
         Caption         =   "CREAT MASK"
         Height          =   375
         Left            =   1440
         TabIndex        =   7
         Top             =   1560
         Width           =   1335
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
         TabIndex        =   6
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
         TabIndex        =   5
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Width [um]"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Offset [um]"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Diameter [um]"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "LAYOUT"
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   3000
      Begin VB.TextBox Text4 
         Height          =   270
         Left            =   1440
         TabIndex        =   16
         Text            =   "0.1"
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   270
         Left            =   1440
         TabIndex        =   14
         Text            =   "0.7"
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   1440
         TabIndex        =   10
         Text            =   "0.03"
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Preview 
         Caption         =   "PREVIEW"
         Height          =   375
         Left            =   1440
         TabIndex        =   8
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Offset [um]"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Length [um]"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Font 
         Caption         =   "Font [um]"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.TextBox TextPath 
      Height          =   420
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6120
   End
   Begin VB.Menu File 
      Caption         =   "File"
      WindowList      =   -1  'True
      Begin VB.Menu LoadExcel 
         Caption         =   "Load Excel"
      End
   End
   Begin VB.Menu About 
      Caption         =   "About"
      Begin VB.Menu AutoLayout 
         Caption         =   "About AutoLayout"
      End
   End
End
Attribute VB_Name = "EntranceForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gActiveDoc As IAcadDocument
Dim gAcadApplication As IAcadApplication
Private Sub AutoLayout_Click()
    Dim newform As New AboutForm
    newform.Show
End Sub
Private Sub LoadExcel_Click()
   On Error GoTo ErrHandler
    CommonDialog1.Filter = "excelfile (*.xls)|*.xls|"
    CommonDialog1.ShowOpen
    TextPath.Text = CommonDialog1.FileName
    Exit Sub
ErrHandler:
    Exit Sub
End Sub
Private Sub Preview_Click()
    Dim newform As New ProbeAngleForm
    newform.Show
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
    Set excelApp = M_CreateExcel(TextPath.Text)
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
    Set excelApp = M_CreateExcel(TextPath.Text)
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
