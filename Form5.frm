VERSION 5.00
Begin VB.Form LayerForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Layer Assignment"
   ClientHeight    =   3540
   ClientLeft      =   2745
   ClientTop       =   1350
   ClientWidth     =   3570
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   3570
   Begin VB.CommandButton ClearCmd 
      Caption         =   "Clear"
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
      Left            =   1560
      TabIndex        =   26
      Top             =   3080
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   840
      TabIndex        =   25
      Top             =   1050
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "B��"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   24
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "�� R"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   23
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "L ��"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   22
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "T��"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   21
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Assign"
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
      Top             =   3080
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
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
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   2400
      TabIndex        =   0
      Top             =   0
      Width           =   1095
      Begin VB.Label Label18 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  9"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   3120
         Width           =   375
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 18"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   600
         TabIndex        =   19
         Top             =   3120
         Width           =   375
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 3"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF00FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  7"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   2760
         Width           =   375
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 10"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   600
         TabIndex        =   11
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 11"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   600
         TabIndex        =   10
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 12"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   600
         TabIndex        =   9
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 13"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   600
         TabIndex        =   8
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 14"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   600
         TabIndex        =   7
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 15"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 16"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   600
         TabIndex        =   5
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00404080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 17"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   2760
         Width           =   375
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   375
      End
   End
End
Attribute VB_Name = "LayerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim layerCount As Integer
Dim layerOrder(0 To 20) As Integer
Dim direction As Integer
'order dircetion
'1: X from low to high
'2: X from high to low
'3: Y from low to high
'4: Y from hight to low

Private Sub ClearCmd_Click()
    Text1.Text = ""
    layerCount = 0
End Sub

Private Sub Command1_Click()
'layerArray() As Integer, arrayLength As Integer, orderDirection As Integer
    Call ReorderLayer(layerOrder, layerCount, direction)
    Call M_RedrawPicutreBox
    Call LayerForm.Hide
End Sub

Private Sub Command2_Click()
    LayerForm.Text2.Text = "T"
    direction = 4
End Sub

Private Sub Command3_Click()
    LayerForm.Text2.Text = "L"
    direction = 1
End Sub

Private Sub Command4_Click()
    LayerForm.Text2.Text = "R"
    direction = 2
End Sub

Private Sub Command5_Click()
    LayerForm.Text2.Text = "B"
    direction = 3
End Sub

Private Sub Form_Load()
    layerCount = 0
    direction = 1
End Sub
Private Sub AddLayerOrder(layerNo As Integer)
    If (layerCount < 20) Then
        layerOrder(layerCount) = layerNo
        layerCount = layerCount + 1
    End If
End Sub

Private Sub Label1_Click()
    LayerForm.Text1.Text = Text1.Text & "1,"
    Call AddLayerOrder(1)
End Sub

Private Sub Label10_Click()
    LayerForm.Text1.Text = Text1.Text & "8,"
    Call AddLayerOrder(8)
End Sub

Private Sub Label11_Click()
    LayerForm.Text1.Text = Text1.Text & "7,"
    Call AddLayerOrder(7)
End Sub

Private Sub Label12_Click()
    LayerForm.Text1.Text = Text1.Text & "6,"
    Call AddLayerOrder(6)
End Sub

Private Sub Label13_Click()
    LayerForm.Text1.Text = Text1.Text & "5,"
    Call AddLayerOrder(5)
End Sub

Private Sub Label14_Click()
    LayerForm.Text1.Text = Text1.Text & "4,"
    Call AddLayerOrder(4)
End Sub

Private Sub Label15_Click()
    LayerForm.Text1.Text = Text1.Text & "3,"
    Call AddLayerOrder(3)
End Sub

Private Sub Label16_Click()
    LayerForm.Text1.Text = Text1.Text & "2,"
    Call AddLayerOrder(2)
End Sub

Private Sub Label17_Click()
    LayerForm.Text1.Text = Text1.Text & "18,"
    Call AddLayerOrder(18)
End Sub

Private Sub Label18_Click()
    LayerForm.Text1.Text = Text1.Text & "9,"
    Call AddLayerOrder(9)
End Sub

Private Sub Label2_Click()
    LayerForm.Text1.Text = Text1.Text & "17,"
    Call AddLayerOrder(17)
End Sub

Private Sub Label3_Click()
    LayerForm.Text1.Text = Text1.Text & "16,"
    Call AddLayerOrder(16)
End Sub

Private Sub Label4_Click()
    LayerForm.Text1.Text = Text1.Text & "15,"
    Call AddLayerOrder(15)
End Sub

Private Sub Label5_Click()
    LayerForm.Text1.Text = Text1.Text & "14,"
    Call AddLayerOrder(14)
End Sub

Private Sub Label6_Click()
    LayerForm.Text1.Text = Text1.Text & "13,"
    Call AddLayerOrder(13)
End Sub

Private Sub Label7_Click()
    LayerForm.Text1.Text = Text1.Text & "12,"
    Call AddLayerOrder(12)
End Sub

Private Sub Label8_Click()
    LayerForm.Text1.Text = Text1.Text & "11,"
    Call AddLayerOrder(11)
End Sub

Private Sub Label9_Click()
    LayerForm.Text1.Text = Text1.Text & "10,"
    Call AddLayerOrder(10)
End Sub
