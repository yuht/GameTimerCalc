VERSION 5.00
Begin VB.Form Form1 
<<<<<<< HEAD
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "刷怪时间计算器    by滴滴地滴滴  BiuBiu"
   ClientHeight    =   4215
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6540
=======
   BorderStyle     =   1  'Fixed Single
   Caption         =   "刷怪时间提示  by滴滴地滴滴 BiuBiu"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6510
>>>>>>> 47e1e555c5bca79f46ab5508f1a6d1bfd0edf091
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
   ScaleHeight     =   5280
   ScaleWidth      =   6510
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame2 
      BackColor       =   &H80000005&
      Caption         =   "刷怪时间"
      Height          =   5235
      Left            =   45
      TabIndex        =   4
      Top             =   15
      Width           =   3705
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   3105
         Top             =   90
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         Height          =   4515
         ItemData        =   "Form1.frx":030A
         Left            =   90
         List            =   "Form1.frx":030C
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   630
         Width           =   3540
      End
      Begin VB.Label Label1 
<<<<<<< HEAD
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
=======
         BackStyle       =   0  'Transparent
>>>>>>> 47e1e555c5bca79f46ab5508f1a6d1bfd0edf091
         Caption         =   "当前时间："
         Height          =   195
         Left            =   135
         TabIndex        =   6
         Top             =   315
         Width           =   3510
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      Caption         =   "怪物和刷新时间管理"
      Height          =   5235
      Left            =   3780
      TabIndex        =   0
      Top             =   15
      Width           =   2670
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "删除"
         Height          =   870
         Left            =   2115
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2070
         Width           =   465
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "添加"
         Height          =   870
         Left            =   2115
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   855
         Width           =   465
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   4905
         Left            =   90
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   260
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
    Name As String
    Dt As String
    refDt As String
    
End Type

Private Type typeListDetial
    Name As String
    NextDispTime As String
    DispFlag As Boolean
End Type


Dim JuanJuanRef() As typeJuanJuanRef
Dim ListDetials() As typeListDetial

Dim oldInt As Integer
Dim RefNow As Boolean

Dim NeedRefreshList2 As Boolean

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
    Call TimeRef
End Sub

Private Sub Form_Load()
    Dim DCTT As String
    List1.Clear
    List2.Clear
    
    Call refList
    
    NeedRefreshList2 = False
     
    
    Me.Caption = Me.Caption & "   Ver:" & App.Major & "." & App.Minor & "." & App.Revision
    Call TimeRef
    Call Timer1_Timer
End Sub



Public Function refList()
    
    Dim i As Integer
    Dim j As Integer
    Dim SectionNames As String
    Dim ArraySectionNames() As String
    Dim UbndASN As Integer
    Dim testbyte() As Byte
    List1.Clear
    
    SectionNames = GetSectionNames()
    SectionNames = Replace(SectionNames, "Setting" & vbNullChar, "")
    SectionNames = Replace(SectionNames, vbNullChar & vbNullChar, "")
    
     
    If Len(SectionNames) = 0 Then
        Exit Function
    End If
    
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
    
    ReDim ListDetials(UbndASN, 3)
    
    ReDim JuanJuanRef(UbndASN)
    
    For i = 0 To UbndASN
        List1.AddItem ArraySectionNames(i)
        JuanJuanRef(i).Name = ArraySectionNames(i)
        
        JuanJuanRef(i).Dt = GetFromInI(JuanJuanRef(i).Name, "Dt")
        If JuanJuanRef(i).Dt = "" Then
            JuanJuanRef(i).Dt = Format(Now, "YYYY/MM/DD HH:MM:SS")
            Call PutToInI(JuanJuanRef(i).Name, "Dt", JuanJuanRef(i).Dt)
        End If

        
        JuanJuanRef(i).refDt = GetFromInI(JuanJuanRef(i).Name, "refDt")
        If JuanJuanRef(i).refDt = "" Then
            JuanJuanRef(i).refDt = TimeSerial(Hour(Now), Minute(Now), Second(Now))
            Call PutToInI(JuanJuanRef(i).Name, "refDt", JuanJuanRef(i).refDt)
        End If
    
    
        For j = 0 To UBound(ListDetials, 2)
            ListDetials(i, j).Name = JuanJuanRef(i).Name
        Next
    
    Next
            
    Timer1.Enabled = True
    
End Function

Public Function TimeRef()
    
<<<<<<< HEAD
On Error Resume Next
    Dim Errmsg  As String
=======
<<<<<<< HEAD
    Dim i As Double
    Dim j As Double
    Dim k As Double
    Dim seci As Double
=======
>>>>>>> 47e1e555c5bca79f46ab5508f1a6d1bfd0edf091
    Dim i As Integer
    Dim j As Integer
    Dim k
    Dim seci
>>>>>>> origin/master
    
    Dim SecDtj
    Dim SecNow
    Dim IntSecNowDivDtj
    
    Dim Dtj As String
<<<<<<< HEAD
    Dim Dtjold As String
    Dim DtNow As String
    Dim subSec As Double
=======
    
    
>>>>>>> 47e1e555c5bca79f46ab5508f1a6d1bfd0edf091
    
    Errmsg = ""
    List2.Clear
    If List1.ListCount = 0 Then
        Exit Function
    End If
    
    Timer1.Enabled = False
    
    j = UBound(JuanJuanRef)
    For i = 0 To j
        Err.Clear
        
        '把刷新间隔时间转成秒
        seci = TimeGetSeconds(JuanJuanRef(i).refDt)
        '上次刷新时间
        Dtj = JuanJuanRef(i).Dt
        
        '刷新时间不为零
        If seci > 0 Then
'            '获取时间差
            SecDtj = TimeGetUTCSeconds(JuanJuanRef(i).Dt)
            SecNow = TimeGetUTCSeconds(Now)
            
            IntSecNowDivDtj = Fix((SecNow - SecDtj) / seci)
            
            Dtj = DateAdd("s", IntSecNowDivDtj * seci, Dtj)
            Call PutToInI(JuanJuanRef(i).Name, "Dt", Dtj)
        End If
<<<<<<< HEAD
        
        List2.AddItem Format(Dtj, "YYYY/MM/DD HH:MM:SS") & " - " & JuanJuanRef(i).name
        If Err Then
            If Len(Errmsg) Then
                Errmsg = Errmsg + vbCrLf + vbCrLf
            End If
            Errmsg = Errmsg + "怪物名称：" + JuanJuanRef(i).name + vbCrLf + "刷怪时间：" + JuanJuanRef(i).Dt + vbCrLf + "刷新间隔：" + JuanJuanRef(i).refDt
        End If
        DoEvents
    Next
    
    If Err Then
        Timer1.Enabled = False
        MsgBox Errmsg + vbCrLf + vbCrLf + "修正参数后关闭程序重新运行", vbCritical + vbOKOnly
    End If
    
=======
        For k = 0 To UBound(ListDetials, 2)
            ListDetials(i, k).NextDispTime = TimeFormat(DateAdd("s", k * seci, Dtj))
            ListDetials(i, k).DispFlag = False
            'List2.AddItem ListDetials(i, k).NextDispTime & " -    " & ListDetials(i, k).Name
            DoEvents
        Next
        ListDetials(i, 0).DispFlag = True
        DoEvents
    Next
    
    Timer1.Enabled = True
     
>>>>>>> 47e1e555c5bca79f46ab5508f1a6d1bfd0edf091
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
        Form2.Show
    End If
    Form2.Command1.Caption = "修改"
End Sub
 


Private Function GetShortestSeconds()
    Dim i As Integer
    Dim shortestSEC
    Dim dtSec
    shortestSEC = 2000
    dtSec = 0
    
    
    If Not List1.ListCount = 0 Then
        For i = 0 To UBound(JuanJuanRef)
            dtSec = TimeGetSeconds(JuanJuanRef(i).refDt)
            If shortestSEC > dtSec Then
                shortestSEC = dtSec
            End If
        Next
    End If
    
    GetShortestSeconds = Format(shortestSEC / 60, "0.00")
End Function

Private Sub Timer1_Timer()
    Dim i As Integer
    Dim j As Integer
    Label1.Caption = TimeFormat(Now) & " - 当前时间"
    
    If List1.ListCount = 0 Then
        Exit Sub
    End If
    
    Call refListDetials
    
    If NeedRefreshList2 = True Or List2.ListCount = 0 Then
        List2.Clear
        For i = 0 To UBound(ListDetials, 1)
            For j = 0 To UBound(ListDetials, 2)
                List2.AddItem ListDetials(i, j).NextDispTime & " - " & IIf(ListDetials(i, j).DispFlag = True, "√", "  ") & " " & ListDetials(i, j).Name
            Next
        Next
    End If
End Sub

Function refListDetials()
    Dim i As Integer
    Dim j As Integer
    For i = 0 To UBound(ListDetials, 1)
        While (DateDiff("s", ListDetials(i, 1).NextDispTime, Now) >= 0) And (TimeGetSeconds(JuanJuanRef(i).refDt) <> 0)
            For j = 1 To UBound(ListDetials, 2)
                ListDetials(i, j - 1) = ListDetials(i, j)
                DoEvents
            Next
            ListDetials(i, UBound(ListDetials, 2)).NextDispTime = TimeFormat(DateAdd("s", TimeGetSeconds(JuanJuanRef(i).refDt), ListDetials(i, UBound(ListDetials, 2)).NextDispTime))
            ListDetials(i, 0).DispFlag = True
            NeedRefreshList2 = True
            DoEvents
        Wend
        
         If TimeGetSeconds(JuanJuanRef(i).refDt) = 0 Then
             For j = 0 To UBound(ListDetials, 2)
                If DateDiff("s", ListDetials(i, j).NextDispTime, Now) >= 0 Then
                    ListDetials(i, j).DispFlag = True
                End If
                DoEvents
            Next
        End If
        
        DoEvents
    Next
End Function

Function TimeGetUTCSeconds(time As String)
    TimeGetUTCSeconds = DateDiff("s", "1970/01/01 00:00:00", time)
End Function

Function TimeFormat(time As String) As String
    TimeFormat = Format(Trim$(time), "YYYY/MM/DD HH:MM:SS")
End Function

Function TimeGetSeconds(time As String) As Integer
    TimeGetSeconds = DateDiff("s", "1970/01/01 00:00:00", "1970/01/01 " & time)
End Function
