VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   3690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3975
   LinkTopic       =   "Form5"
   ScaleHeight     =   3690
   ScaleWidth      =   3975
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   960
      TabIndex        =   25
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "B"
      Height          =   495
      Left            =   1080
      TabIndex        =   24
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "R"
      Height          =   495
      Left            =   1920
      TabIndex        =   23
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "L"
      Height          =   495
      Left            =   240
      TabIndex        =   22
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "T"
      Height          =   495
      Left            =   1080
      TabIndex        =   21
      Top             =   120
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
      Left            =   240
      TabIndex        =   2
      Top             =   3120
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   1095
      Begin VB.Label Label18 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 9"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   3120
         Width           =   375
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 18"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   600
         TabIndex        =   19
         Top             =   3120
         Width           =   375
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 2"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 4"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 5"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 6"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 7"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 8"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   2760
         Width           =   375
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 10"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   600
         TabIndex        =   11
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 11"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   600
         TabIndex        =   10
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 12"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   600
         TabIndex        =   9
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 13"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   600
         TabIndex        =   8
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 14"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   600
         TabIndex        =   7
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 15"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 16"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   600
         TabIndex        =   5
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 17"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   2760
         Width           =   375
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 1"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   375
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Form5.Text2.Text = T
End Sub

Private Sub Command3_Click()
Form5.Text2.Text = l
End Sub

Private Sub Command4_Click()
Form5.Text2.Text = R
End Sub

Private Sub Command5_Click()
Form5.Text2.Text = b
End Sub

Private Sub Label1_Click()
Form5.Text1.Text = Text1.Text & ",1"
End Sub

Private Sub Label10_Click()
Form5.Text1.Text = Text1.Text & ",8"
End Sub

Private Sub Label11_Click()
Form5.Text1.Text = Text1.Text & ",7"
End Sub

Private Sub Label12_Click()
Form5.Text1.Text = Text1.Text & ",6"
End Sub

Private Sub Label13_Click()
Form5.Text1.Text = Text1.Text & ",5"
End Sub

Private Sub Label14_Click()
Form5.Text1.Text = Text1.Text & ",4"
End Sub

Private Sub Label15_Click()
Form5.Text1.Text = Text1.Text & ",3"
End Sub

Private Sub Label16_Click()
Form5.Text1.Text = Text1.Text & ",2"
End Sub

Private Sub Label17_Click()
Form5.Text1.Text = Text1.Text & ",18"
End Sub

Private Sub Label18_Click()
Form5.Text1.Text = Text1.Text & ",9"
End Sub

Private Sub Label2_Click()
Form5.Text1.Text = Text1.Text & ",17"
End Sub

Private Sub Label3_Click()
Form5.Text1.Text = Text1.Text & ",16"
End Sub

Private Sub Label4_Click()
Form5.Text1.Text = Text1.Text & ",15"
End Sub

Private Sub Label5_Click()
Form5.Text1.Text = Text1.Text & ",14"
End Sub

Private Sub Label6_Click()
Form5.Text1.Text = Text1.Text & ",13"
End Sub

Private Sub Label7_Click()
Form5.Text1.Text = Text1.Text & ",12"
End Sub

Private Sub Label8_Click()
Form5.Text1.Text = Text1.Text & ",11"
End Sub

Private Sub Label9_Click()
Form5.Text1.Text = Text1.Text & ",10"
End Sub
