VERSION 5.00
Begin VB.Form SectionForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Section Drawing"
   ClientHeight    =   2535
   ClientLeft      =   2805
   ClientTop       =   1410
   ClientWidth     =   3585
   Icon            =   "SectionForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   3585
   Begin VB.Frame Frame3 
      Caption         =   "Needle Assembly"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3360
      Begin VB.CommandButton Section 
         Caption         =   "SECTION DRAWING"
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
         Left            =   120
         TabIndex        =   2
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox Text14 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         TabIndex        =   1
         Text            =   "94"
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "Theta"
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
         TabIndex        =   4
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Browse Excel of Needle Force First "
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
         TabIndex        =   3
         Top             =   360
         Width           =   3135
      End
   End
End
Attribute VB_Name = "SectionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
