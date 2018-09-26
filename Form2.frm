VERSION 5.00
Begin VB.Form Form2 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "怪物和刷新时间管理"
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5265
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "关闭"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4185
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   945
      Width           =   990
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "添加"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4185
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   517
      Width           =   990
   End
   Begin VB.TextBox txtText_name 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1080
      TabIndex        =   0
      Top             =   60
      Width           =   2670
   End
   Begin VB.TextBox txtText2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   1080
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   510
      Width           =   645
   End
   Begin VB.TextBox txtText2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   2160
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   510
      Width           =   645
   End
   Begin VB.TextBox txtText2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   2
      Left            =   3105
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   510
      Width           =   645
   End
   Begin VB.CommandButton cmd 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "当前时间"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4185
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   90
      Width           =   990
   End
   Begin VB.TextBox txtText2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   3
      Left            =   1080
      TabIndex        =   5
      Top             =   975
      Width           =   645
   End
   Begin VB.TextBox txtText2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   4
      Left            =   2160
      TabIndex        =   6
      Top             =   975
      Width           =   645
   End
   Begin VB.TextBox txtText2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   5
      Left            =   3105
      TabIndex        =   7
      Text            =   "0"
      Top             =   975
      Width           =   645
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "怪物名称："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   6
      Left            =   45
      TabIndex        =   17
      Top             =   135
      Width           =   1050
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "刷怪时间："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   45
      TabIndex        =   16
      Top             =   585
      Width           =   1050
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   1845
      TabIndex        =   15
      Top             =   585
      Width           =   210
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   2880
      TabIndex        =   14
      Top             =   585
      Width           =   210
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "刷新间隔："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   45
      TabIndex        =   13
      Top             =   1050
      Width           =   1050
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "分"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   4
      Left            =   2880
      TabIndex        =   12
      Top             =   1050
      Width           =   210
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "秒"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   5
      Left            =   3825
      TabIndex        =   11
      Top             =   1050
      Width           =   210
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "小时"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   9
      Left            =   1755
      TabIndex        =   10
      Top             =   1050
      Width           =   420
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ka As Date
        
Private Sub cmd_Click()
    txtText2(0) = Hour(Now)
    txtText2(1) = Minute(Now)
    txtText2(2) = Second(Now)
End Sub

Private Sub Command1_Click()
    Dim Dt As String, refDt As String

    Dim i As Integer
        
    For i = 0 To 5
        If Not IsNumeric(txtText2(i)) Then
            txtText2(i) = 0
        End If
    Next
    
    Dt = Date & " " & TimeSerial(txtText2(0), txtText2(1), txtText2(2))
    refDt = TimeSerial(txtText2(3), txtText2(4), txtText2(5))
    
'    MsgBox txtText_name & vbCrLf & Date & " " & Dt & vbCrLf & refDt
    
    Call PutToInI(txtText_name, "Dt", Dt)
    Call PutToInI(txtText_name, "refDt", refDt)
    
    Call Form1.refList
    
    
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    ka = Date
    txtText2(0) = Hour(Now)
    txtText2(1) = Minute(Now)
    txtText2(2) = Second(Now)
End Sub
