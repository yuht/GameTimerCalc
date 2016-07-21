VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "刷怪时间计算器    by滴滴地滴滴  BiuBiu"
   ClientHeight    =   4095
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7545
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   7545
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox txtText2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   330
      Index           =   6
      Left            =   1215
      TabIndex        =   17
      Text            =   "30"
      Top             =   1365
      Width           =   645
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "开始计算"
      Height          =   405
      Left            =   2700
      TabIndex        =   16
      Top             =   1350
      Width           =   1395
   End
   Begin VB.TextBox txtText1 
      Appearance      =   0  'Flat
      Height          =   1995
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Top             =   2025
      Width           =   7440
   End
   Begin VB.TextBox txtText2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   330
      Index           =   5
      Left            =   2970
      TabIndex        =   13
      Text            =   "0"
      Top             =   735
      Width           =   645
   End
   Begin VB.TextBox txtText2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   330
      Index           =   4
      Left            =   2115
      TabIndex        =   9
      Top             =   735
      Width           =   645
   End
   Begin VB.TextBox txtText2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   330
      Index           =   3
      Left            =   1035
      TabIndex        =   8
      Top             =   735
      Width           =   645
   End
   Begin VB.CommandButton cmd 
      Caption         =   "获取时间"
      Height          =   360
      Left            =   3420
      TabIndex        =   3
      Top             =   90
      Width           =   990
   End
   Begin VB.TextBox txtText2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   330
      Index           =   2
      Left            =   2565
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   105
      Width           =   645
   End
   Begin VB.TextBox txtText2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   330
      Index           =   1
      Left            =   1665
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   105
      Width           =   645
   End
   Begin VB.TextBox txtText2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   330
      Index           =   0
      Left            =   765
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   105
      Width           =   645
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "小时"
      Height          =   195
      Index           =   9
      Left            =   1710
      TabIndex        =   18
      Top             =   810
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "次"
      Height          =   195
      Index           =   7
      Left            =   1935
      TabIndex        =   14
      Top             =   1440
      Width           =   180
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "计算多少次："
      Height          =   195
      Index           =   6
      Left            =   90
      TabIndex        =   12
      Top             =   1440
      Width           =   1080
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "秒"
      Height          =   195
      Index           =   5
      Left            =   3690
      TabIndex        =   11
      Top             =   810
      Width           =   180
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "分"
      Height          =   195
      Index           =   4
      Left            =   2790
      TabIndex        =   10
      Top             =   810
      Width           =   180
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "刷新间隔："
      Height          =   195
      Index           =   3
      Left            =   90
      TabIndex        =   7
      Top             =   810
      Width           =   900
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "："
      Height          =   195
      Index           =   2
      Left            =   2340
      TabIndex        =   6
      Top             =   180
      Width           =   180
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "："
      Height          =   195
      Index           =   1
      Left            =   1440
      TabIndex        =   5
      Top             =   180
      Width           =   180
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "时间："
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   4
      Top             =   180
      Width           =   540
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Click()
    txtText2(0) = Hour(Now)
    txtText2(1) = Minute(Now)
    txtText2(2) = Second(Now)
End Sub

Private Sub cmd2_Click()
    
    Dim i As Integer
    Dim dt
    txtText1 = ""
    Dim times As Integer
    For i = 0 To 6
        If Not IsNumeric(txtText2(i)) Then
            txtText2(i) = 0
        End If
    Next
    
    times = CInt(txtText2(6))
    For i = 1 To times
        dt = TimeSerial(txtText2(0) + txtText2(3) * i, txtText2(1) + txtText2(4) * i, txtText2(2) + txtText2(5) * i)
            
        txtText1 = txtText1 & IIf(i = 1, "", " , ") & Format(dt, "hh:mm:ss")
    
    Next
    
End Sub

Private Sub Form_Load()
    Call cmd_Click
End Sub

