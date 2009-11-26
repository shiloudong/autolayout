VERSION 5.00
Begin VB.Form MaskForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mask DXF"
   ClientHeight    =   2385
   ClientLeft      =   2865
   ClientTop       =   1425
   ClientWidth     =   2580
   Icon            =   "Mask.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   2580
   Begin VB.Frame Frame2 
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
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2400
      Begin VB.CommandButton Command1 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   8
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1440
         TabIndex        =   4
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1440
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton mask 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1440
         TabIndex        =   1
         Text            =   "25"
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Diameter [um]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Offset [um]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Width [um]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   975
      End
   End
End
Attribute VB_Name = "MaskForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

'Creat Mask
Private Sub mask_Click()
    Call MaskForm.Hide
    Dim app As IAcadApplication
    Dim doc2 As IAcadDocument
    Set app = CreateAcad
    Set doc2 = app.Documents.Add
    
    Set excelApp = M_CreateExcel(ProbeAngleForm.CommonDialog1.FileName)
    Set excelsheet = excelApp.ActiveWorkbook.Sheets("sheet1") '当前工作表为sheet1
    
    Dim maskp(0 To 2) As Double
    Dim ed As Double
    Dim id As Double

    id = MaskForm.Text2.Text / 1000
    ed = MaskForm.Text2.Text / 1000 + MaskForm.Text3.Text / 1000

    Dim newCenter(0 To 2) As Double
    newCenter(0) = (M_MaxX + M_MinX) / 2
    newCenter(1) = (M_MaxY + M_MinY) / 2
    newCenter(2) = 0
    
For i = 0 To M_RowCount - 1
    x = M_PointsX(i)
    y = M_PointsY(i)
    Angle = M_Angles(i)
        
    maskp(0) = (M_PointsX(i) - newCenter(0)) + MaskForm.Text1.Text / 1000 * Cos(3.1415926 * Angle / 180)
    maskp(1) = (M_PointsY(i) - newCenter(1)) + MaskForm.Text1.Text / 1000 * Sin(3.1415926 * Angle / 180)
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
    Call doc2.ModelSpace.AddText("Dia=" & Text2.Text, p12, 1.5)
    Call doc2.ModelSpace.AddText("Offset=" & Text1.Text, p13, 1.5)
    doc2.Application.ZoomExtents
    Call excelApp.Workbooks.Close '关闭excel程序
    
End Sub
