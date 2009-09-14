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
      Left            =   5520
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
      TabIndex        =   10
      Top             =   2760
      Width           =   6120
      Begin VB.TextBox Text14 
         Height          =   285
         Left            =   1440
         TabIndex        =   18
         Text            =   "94"
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Section 
         Caption         =   "SECTION DRAWING"
         Height          =   375
         Left            =   3360
         TabIndex        =   11
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Browse Excel of Needle Force First "
         Height          =   255
         Left            =   2760
         TabIndex        =   20
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label14 
         Caption         =   "Theta"
         Height          =   255
         Left            =   600
         TabIndex        =   19
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "MASK"
      Height          =   2055
      Left            =   3240
      TabIndex        =   1
      Top             =   600
      Width           =   3000
      Begin VB.TextBox Text5 
         Height          =   270
         Left            =   1560
         TabIndex        =   17
         Text            =   "25"
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton mask 
         Caption         =   "CREAT MASK"
         Height          =   375
         Left            =   1440
         TabIndex        =   6
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox Text11 
         BeginProperty Font 
            Name            =   "ËÎÌו"
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
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox Text10 
         BeginProperty Font 
            Name            =   "ËÎÌו"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1560
         TabIndex        =   4
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Width [um]"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Offset [um]"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Diameter [um]"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "LAYOUT"
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3000
      Begin VB.TextBox Text4 
         Height          =   270
         Left            =   1440
         TabIndex        =   15
         Text            =   "0.1"
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   270
         Left            =   1440
         TabIndex        =   13
         Text            =   "0.7"
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   1440
         TabIndex        =   9
         Text            =   "0.03"
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Preview 
         Caption         =   "PREVIEW"
         Height          =   375
         Left            =   1440
         TabIndex        =   7
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Offset [um]"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Length [um]"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Font 
         Caption         =   "Font [um]"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   855
      End
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


Private Sub LoadExcel_Click()
   On Error GoTo ErrHandler
    CommonDialog1.Filter = "excelfile (*.xls)|*.xls|"
    CommonDialog1.ShowOpen
    'TextPath.Text = CommonDialog1.FileName
    Call Preview_Click
    Exit Sub
ErrHandler:
    Exit Sub
End Sub
Private Sub Preview_Click()
    Dim newform As New ProbeAngleForm
    newform.Show
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
