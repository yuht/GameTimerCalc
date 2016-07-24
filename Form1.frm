VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "刷怪时间计算器    by滴滴地滴滴  BiuBiu"
   ClientHeight    =   4215
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6540
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   6540
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame2 
      Caption         =   "刷怪时间"
      Height          =   4110
      Left            =   45
      TabIndex        =   4
      Top             =   45
      Width           =   3705
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   3150
         Top             =   225
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         Height          =   3345
         Left            =   45
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   540
         Width           =   3570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "当前时间："
         Height          =   195
         Left            =   90
         TabIndex        =   6
         Top             =   270
         Width           =   900
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "怪物和刷新时间管理"
      Height          =   4110
      Left            =   3825
      TabIndex        =   0
      Top             =   45
      Width           =   2670
      Begin VB.CommandButton Command3 
         Caption         =   "删除"
         Height          =   870
         Left            =   2115
         TabIndex        =   3
         Top             =   2070
         Width           =   465
      End
      Begin VB.CommandButton Command1 
         Caption         =   "添加"
         Height          =   870
         Left            =   2115
         TabIndex        =   2
         Top             =   855
         Width           =   465
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   3735
         Left            =   90
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   270
         Width           =   1950
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type typeJuanJuanRef
    name As String
    Dt As String
    refDt As String
    
End Type
Dim JuanJuanRef() As typeJuanJuanRef
     
Private Sub Command1_Click()
    Form2.Left = Me.Left + Me.Width + 200
    Form2.Top = Me.Top
    Form2.Command1.Caption = "添加"
    Form2.Show
End Sub

 

 

 

Private Sub Command3_Click()
    If List1.SelCount Then
        'MsgBox List1.List(List1.ListIndex)
        Call DelIniSec(List1.List(List1.ListIndex))
        List1.RemoveItem (List1.ListIndex)
    End If
    
    If Form2.Visible Then
        Unload Form2
    End If
    Call refList
End Sub

Private Sub Form_Load()
    List1.Clear
    List2.Clear
    Call refList
End Sub



Public Function refList()
    
    Dim i As Integer
    Dim SectionNames As String
    Dim ArraySectionNames() As String
    Dim UbndASN As Integer
    Dim testbyte() As Byte
    List1.Clear
    
    SectionNames = GetSectionNames()
    SectionNames = Replace(SectionNames, vbNullChar & vbNullChar, "")
    If InStrRev(SectionNames, vbNullChar) = Len(SectionNames) Then
        SectionNames = Left(SectionNames, Len(SectionNames) - 1)
    End If
    ArraySectionNames = Split(SectionNames, vbNullChar)
    testbyte = SectionNames
    
'    'the array aSections now contains section headers
'    MsgBox UBound(ArraySectionNames)
    UbndASN = UBound(ArraySectionNames)
    If (UbndASN) < 0 Then
        Timer1.Enabled = False
        List2.Clear
        Exit Function
    End If
    
    Timer1.Enabled = True
    
    UbndASN = UbndASN
    
    ReDim JuanJuanRef(UbndASN)
    For i = 0 To UbndASN
        List1.AddItem ArraySectionNames(i)
        JuanJuanRef(i).name = ArraySectionNames(i)
        JuanJuanRef(i).Dt = GetFromInI(JuanJuanRef(i).name, "Dt")
        JuanJuanRef(i).refDt = GetFromInI(JuanJuanRef(i).name, "refDt")
    Next
    
    
    
End Function

Private Function TimeRef()
    
    Dim i As Integer
    Dim j As Integer
    Dim seci As Integer
    Dim Dtj As String
    Dim Dtjold As String
    Dim DtNow As String
    Dim subSec As Integer
    
    List2.Clear
    
    j = UBound(JuanJuanRef)
    For i = 0 To j
        '把刷新间隔时间转成秒
        seci = DateDiff("s", "1970/01/01 00:00:00", "1970/01/01 " & JuanJuanRef(i).refDt)
        
        '获取当前时间
        DtNow = Format(Now, "YYYY/MM/DD HH:MM:SS")
        
        '上次刷新时间
        Dtj = JuanJuanRef(i).Dt
        
        '刷新时间不为零
        If seci > 0 Then
            '获取时间差
            
            Dtjold = Dtj
            subSec = DateDiff("s", Dtj, Now)
            While subSec > 0
                Dtjold = Dtj
                Dtj = DateAdd("s", seci, Dtj)
                subSec = DateDiff("s", Dtj, Now)
                DoEvents
            Wend
            Call PutToInI(JuanJuanRef(i).name, "Dt", Dtjold)
        End If
        
        List2.AddItem Format(Dtj, "YYYY/MM/DD HH:MM:SS") & " - " & JuanJuanRef(i).name
        DoEvents
    Next
    
End Function



Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub List1_Click()
    Form2.txtText_name = List1.List(List1.ListIndex)
    Form2.txtText2(0) = Hour(GetFromInI(Form2.txtText_name, "Dt"))
    Form2.txtText2(1) = Minute(GetFromInI(Form2.txtText_name, "Dt"))
    Form2.txtText2(2) = Second(GetFromInI(Form2.txtText_name, "Dt"))
    
    Form2.txtText2(3) = Hour(GetFromInI(Form2.txtText_name, "refDt"))
    Form2.txtText2(4) = Minute(GetFromInI(Form2.txtText_name, "refDt"))
    Form2.txtText2(5) = Second(GetFromInI(Form2.txtText_name, "refDt"))
    
    If Form2.Visible = False Then
        Form2.Left = Me.Left + Me.Width + 200
        Form2.Top = Me.Top
        Form2.Command1.Caption = "修改"
        Form2.Show
    End If
End Sub

Private Sub Timer1_Timer()
    Label1.Caption = Format(Now, "YYYY/MM/DD HH:MM:SS") & " - 当前时间"
    Call TimeRef
End Sub
