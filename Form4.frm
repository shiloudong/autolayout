VERSION 5.00
Begin VB.Form AngleForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Angle"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2775
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   2775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Frame Frame1 
      Caption         =   "Direction"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      Begin VB.CommandButton Command1 
         Caption         =   "Set Probes to  Angle"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   3240
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   1440
         TabIndex        =   10
         Top             =   2400
         Width           =   975
      End
      Begin VB.Frame Frame3 
         Height          =   2055
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2295
         Begin VB.Line Line4 
            X1              =   720
            X2              =   1440
            Y1              =   720
            Y2              =   1440
         End
         Begin VB.Line Line3 
            X1              =   720
            X2              =   1440
            Y1              =   1440
            Y2              =   720
         End
         Begin VB.Line Line2 
            X1              =   720
            X2              =   1560
            Y1              =   1080
            Y2              =   1080
         End
         Begin VB.Line Line1 
            X1              =   1080
            X2              =   1080
            Y1              =   600
            Y2              =   1560
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " 270"
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
            Left            =   840
            TabIndex        =   9
            Top             =   1680
            Width           =   495
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "    0"
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
            Left            =   1680
            TabIndex        =   8
            Top             =   960
            Width           =   495
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " 180"
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
            TabIndex        =   7
            Top             =   960
            Width           =   495
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " 135"
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
            TabIndex        =   6
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " 45"
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
            Left            =   1680
            TabIndex        =   5
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " 225"
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
            TabIndex        =   4
            Top             =   1560
            Width           =   495
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " 315"
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
            Left            =   1680
            TabIndex        =   3
            Top             =   1560
            Width           =   495
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "  90"
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
            Left            =   840
            TabIndex        =   2
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Label Label9 
         Caption         =   "Rotaton Angle:"
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
         TabIndex        =   11
         Top             =   2400
         Width           =   1215
      End
   End
End
Attribute VB_Name = "AngleForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Angle As Double
Private Sub Command1_Click()
    If (Form4.Text1 <> "") Then
        Angle = Form4.Text1
        Call SetSelectedAngle(Angle)
    End If
End Sub

Private Sub Label1_Click()
    Angle = 90
    Form4.Text1 = Angle
End Sub

Private Sub Label2_Click()
    Angle = 315
    Form4.Text1 = Angle
End Sub

Private Sub Label3_Click()
    Angle = 225
    Form4.Text1 = Angle
End Sub

Private Sub Label4_Click()
    Angle = 45
    Form4.Text1 = Angle
End Sub

Private Sub Label5_Click()
    Angle = 135
    Form4.Text1 = Angle
End Sub

Private Sub Label6_Click()
    Angle = 180
    Form4.Text1 = Angle
End Sub

Private Sub Label7_Click()
    Angle = 0
    Form4.Text1 = Angle
End Sub

Private Sub Label8_Click()
    Angle = 270
    Form4.Text1 = Angle
End Sub
