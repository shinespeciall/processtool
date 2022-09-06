VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ProcessTools(1.0)"
   ClientHeight    =   5388
   ClientLeft      =   6816
   ClientTop       =   4428
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5388
   ScaleWidth      =   9420
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command15 
      Caption         =   "保存记录"
      Height          =   255
      Left            =   4440
      TabIndex        =   22
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton Command14 
      Caption         =   "关闭监视器"
      Height          =   255
      Left            =   4440
      TabIndex        =   21
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4440
      Top             =   4200
   End
   Begin VB.CommandButton Command13 
      Caption         =   "启动监视器"
      Height          =   255
      Left            =   4440
      TabIndex        =   20
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command10 
      Caption         =   "转到行"
      Height          =   255
      Left            =   4440
      TabIndex        =   15
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "导航栏(右键单击导航，双击删除)"
      Height          =   4575
      Left            =   5760
      TabIndex        =   12
      Top             =   120
      Width           =   3615
      Begin VB.CommandButton Command12 
         Caption         =   "清空"
         Height          =   255
         Left            =   2400
         TabIndex        =   19
         Top             =   4080
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   3720
         Width           =   3375
      End
      Begin VB.CommandButton Command9 
         Caption         =   "清空列表"
         Height          =   255
         Left            =   2400
         TabIndex        =   14
         Top             =   3480
         Width           =   1095
      End
      Begin VB.ListBox List2 
         Height          =   3108
         ItemData        =   "ProcessTool.frx":0000
         Left            =   120
         List            =   "ProcessTool.frx":0002
         TabIndex        =   13
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "取消前置"
      Height          =   255
      Left            =   4440
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "窗体前置"
      Height          =   255
      Left            =   4440
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   4008
      ItemData        =   "ProcessTool.frx":0004
      Left            =   0
      List            =   "ProcessTool.frx":0006
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   0
      Width           =   4215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "计算器(无用)"
      Height          =   255
      Left            =   4440
      TabIndex        =   11
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "杀线程"
      Height          =   255
      Left            =   4440
      TabIndex        =   7
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "杀进程"
      Height          =   255
      Left            =   4440
      TabIndex        =   6
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "搜索单项"
      Height          =   255
      Left            =   4440
      TabIndex        =   8
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "搜索全部"
      Height          =   255
      Left            =   4440
      TabIndex        =   5
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command11 
      Caption         =   "刷新进程"
      Height          =   255
      Left            =   4440
      TabIndex        =   17
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "刷新进线程"
      Height          =   255
      Left            =   4440
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "线程总数："
      Height          =   615
      Left            =   3120
      TabIndex        =   16
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "总行数："
      Height          =   615
      Left            =   2040
      TabIndex        =   10
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "所在行数：1"
      Height          =   615
      Left            =   1080
      TabIndex        =   9
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "进程数："
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   4080
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function TerminateThread Lib "kernel32" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long
Private Declare Function OpenThread Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwThreadId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function NtTerminateProcess Lib "ntdll" (ByVal hProc As Long, ByVal ExitCode As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function ZwQuerySystemInformation Lib "ntdll" (ByVal SystemInformationClass As Long, SystemInformation As Any, ByVal SystemInformationLength As Long, ReturnLength As Long) As Long
Private Declare Sub RtlMoveMemory Lib "ntdll" (Target As Any, ByVal pSource As Long, ByVal Length As Long)
Private Declare Sub GetMem4 Lib "msvbvm60" (ByVal Address As Long, Value As Any)
Private Declare Function GetVersion Lib "kernel32" () As Long

Private Const THREAD_ALL_ACCESS = &H1F03F
Private Const LB_SETHORIZONTALEXTENT = &H194

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Type UNICODE_STRING
Length As Integer
MaxLength As Integer
pszImageName As Long
End Type

Private Type CLIENT_ID
ProcessId As Long
ThreadId As Long
End Type

Private Type SYSTEM_PROCESS_INFORMATION 'XP but 2k
NextEntryOffset As Long
NumberOfThreads As Long
Reserved1(47) As Byte
ImageName As UNICODE_STRING
Reserved2 As Long
UniqueProcessId As Long
HandleCount As Long
Reserved3 As Long
Reserved4(10) As Long
PeakPagefileUsage As Long
PrivatePageCount As Long
Reserved5(11) As Long
End Type

Private Type SYSTEM_THREAD 'XP but 2k
Reserved(8) As Long
'StartAddress As Long
ClientId As CLIENT_ID
Priority As Long
BasePriority As Long
ContextSwitchCount As Long
State As Long
WaitReason As Long
End Type

Private Const SystemProcessAndThreadInformation As Long = 5
Private Const szProcessId  As String = "进程Id:"
Private Const szThreadCount As String = " > 线程数:"
Private Const szProcessName As String = " > 进程名:"
Private Const szThreadId As String = "     线程Id:0x"
'Private Const szThreadState As String = " >> Thread State:"

Sub findfromlist1(findstr2 As String)
Dim i2 As Long 'dim guanbijishu As Integer
For i2 = 0 To List1.ListCount - 1
List1.Selected(i2) = False
Next
If findstr2 = "" Then
findstr2 = InputBox("输入搜索项（这里查找项，区分大小写）", "程序提示")
End If
If findstr2 = "" Then
Exit Sub
End If
For i2 = 0 To List1.ListCount - 1
Label2.Caption = "所在行数：" & i2 + 1
List1.ListIndex = i2
If InStr(List1.List(i2), findstr2) <> 0 Then
MsgBox "恭喜你，找到了项————" & List1.List(i2) & "行数：" & i2 + 1
'guanbijishu = MsgBox("恭喜你，找到了项————" & List1.List(i2) & "行数：" & i2 + 1, vbOKCancel)
'If guanbijishu = 2 Then
'Exit Sub
'End If
List1.ListIndex = i2
List1.SetFocus
List1.Selected(i2) = True
Exit For
End If
Next
End Sub

Sub useshuipingscoll(listfind As ListBox)
Dim max As Long, f As Font, i As Integer
Me.ScaleMode = vbPixels
Set f = Me.Font
Set Me.Font = listfind.Font
With listfind
For i = 0 To .ListCount
If Me.TextWidth(.List(i)) > max Then
max = Me.TextWidth(.List(i))
End If
Next
End With
max = max + 10
Set Me.Font = f
SendMessage listfind.hwnd, LB_SETHORIZONTALEXTENT, max, ByVal 0&
Set f = Nothing
End Sub

Public Sub EnumerateProcessThread(ByVal lst As ListBox) '枚举进程线程,添加到列表
lst.Clear
Dim bfBuffer() As Byte, BufferSize As Long, szInfo As String, jishuthread As Long
Dim szImageName As String, CurPtr As Long, ThreadPtr As Long, ThreadIndex As Long, ThreadId As Long
Dim SystemProcess As SYSTEM_PROCESS_INFORMATION ', ThreadOfProcess() As SYSTEM_THREAD
ZwQuerySystemInformation SystemProcessAndThreadInformation, ByVal 0&, 0, BufferSize
ReDim bfBuffer(BufferSize) As Byte
ZwQuerySystemInformation SystemProcessAndThreadInformation, bfBuffer(0), BufferSize, ByVal 0&
CurPtr = VarPtr(bfBuffer(0)) '当前指针指向第一个进程信息
Dim jishu As Long
Do
jishu = jishu + 1
RtlMoveMemory SystemProcess, CurPtr, Len(SystemProcess)

'####################################################

'szInfo = szProcessId & CStr(SystemProcess.UniqueProcessId)'第一处
szInfo = szProcessId & CStr(SystemProcess.UniqueProcessId) & "(Hex:" & DEC_to_HEX(CStr(SystemProcess.UniqueProcessId)) & ")"
'####################################################

If SystemProcess.ImageName.Length > 0 Then
szImageName = Space(SystemProcess.ImageName.Length)
RtlMoveMemory ByVal StrPtr(szImageName), SystemProcess.ImageName.pszImageName, SystemProcess.ImageName.Length
szInfo = szInfo & szProcessName & Trim(szImageName)
Else
szInfo = szInfo & szProcessName & "<Null>"
End If
szInfo = szInfo & szThreadCount & CStr(SystemProcess.NumberOfThreads)
lst.AddItem szInfo
ThreadPtr = CurPtr + Len(SystemProcess) '线程列表指针指针
'**************************************用于重新枚举系统进程线程,等到加载列表框变色API后使用
'ReDim ThreadOfProcess(SystemProcess.NumberOfThreads - 1) As SYSTEM_THREAD
'RtlMoveMemory ThreadOfProcess(0), ThreadPtr, Len(ThreadOfProcess(0)) * SystemProcess.NumberOfThreads
'For ThreadIndex = 0 To SystemProcess.NumberOfThreads - 1
 '    With ThreadOfProcess(ThreadIndex)
  '       szInfo = szThreadId & Hex(.ClientId.ThreadId)
  '       lst.AddItem szInfo
 '    End With
'Next ThreadIndex
'**************************************
ThreadIndex = 0
Do
jishuthread = jishuthread + 1
GetMem4 ThreadPtr + 40, ThreadId
lst.AddItem szThreadId & Hex(ThreadId)
ThreadPtr = ThreadPtr + 64
ThreadIndex = ThreadIndex + 1
DoEvents
Loop Until ThreadIndex = SystemProcess.NumberOfThreads
CurPtr = CurPtr + SystemProcess.NextEntryOffset '指针转移到下一个指针
Loop Until SystemProcess.NextEntryOffset = 0
Erase bfBuffer
Label1.Caption = "进程数：" & jishu
Label4.Caption = "线程总数： " & jishuthread
End Sub

Private Sub Command1_Click()
Dim jishuqi1 As Long
Call EnumerateProcessThread(List1)
Call useshuipingscoll(List1)
jishuqi1 = List1.ListCount
Label3.Caption = "总行数：" & jishuqi1
End Sub

Private Sub Command10_Click() '暂时没什么卵用
On Error Resume Next
Dim i4 As Integer, i5 As Long
For i5 = 0 To List1.ListCount - 1
List1.Selected(i5) = False
Next
i4 = InputBox("输入行数", "程序提示", 1)
If i4 <= 0 Then
Exit Sub
End If
List1.ListIndex = i4 - 1
List1.SetFocus
List1.Selected(i4 - 1) = True
End Sub

Sub logout()
Unload Me
End
End Sub

Private Sub Command11_Click()
Dim numjishu As Long
numjishu = 0
List1.Clear
Dim bfBuffer() As Byte, BufferSize As Long, szInfo As String
Dim szImageName As String, CurPtr As Long
Dim SystemProcess As SYSTEM_PROCESS_INFORMATION
ZwQuerySystemInformation SystemProcessAndThreadInformation, ByVal 0&, 0, BufferSize
ReDim bfBuffer(BufferSize) As Byte
ZwQuerySystemInformation SystemProcessAndThreadInformation, bfBuffer(0), BufferSize, ByVal 0&
CurPtr = VarPtr(bfBuffer(0)) '当前指针指向第一个进程信息
Dim jishu As Long
Do
jishu = jishu + 1
RtlMoveMemory SystemProcess, CurPtr, Len(SystemProcess)

'####################################################

'szInfo = szProcessId & CStr(SystemProcess.UniqueProcessId)'第二处
szInfo = szProcessId & CStr(SystemProcess.UniqueProcessId) & "(Hex:" & DEC_to_HEX(CStr(SystemProcess.UniqueProcessId)) & ")"
'####################################################

If SystemProcess.ImageName.Length > 0 Then
szImageName = Space(SystemProcess.ImageName.Length)
RtlMoveMemory ByVal StrPtr(szImageName), SystemProcess.ImageName.pszImageName, SystemProcess.ImageName.Length
szInfo = szInfo & szProcessName & Trim(szImageName)
Else
szInfo = szInfo & szProcessName & "<Null>"
End If
szInfo = szInfo & szThreadCount & CStr(SystemProcess.NumberOfThreads)
List1.AddItem szInfo
numjishu = numjishu + Val(CStr(SystemProcess.NumberOfThreads))
CurPtr = CurPtr + SystemProcess.NextEntryOffset '指针转移到下一个指针
Loop Until SystemProcess.NextEntryOffset = 0
Erase bfBuffer
Label1.Caption = "进程数：" & jishu
Call useshuipingscoll(List1)
Dim jishuqi1 As Long
jishuqi1 = List1.ListCount
Label3.Caption = "总行数：" & jishuqi1
Label4.Caption = "线程总数：" + CStr(numjishu)
End Sub

Private Sub Command12_Click()
Text1.Text = ""
End Sub

Private Sub Command13_Click()
Timer1.Enabled = True
End Sub

Private Sub Command14_Click()
Timer1.Enabled = False
End Sub

Private Sub Command15_Click()
Dim Path As String, num As Long
Path = IIf(Len(App.Path) > 3, App.Path & "\", App.Path) '自己的路径
num = 1
Do
If Dir(Path & num & ".txt") = "" Then  '检测是否不存在
Exit Do
End If
num = num + 1
Loop

If List1.ListCount = 0 Then
MsgBox "list中不存在数据！"
Exit Sub
End If
Open Path & num & ".txt" For Output As #1 '创建txt
'****************************************
Dim i As Long
For i = 0 To List1.ListCount - 1
List1.Selected(i) = False
Next

For i = 0 To List1.ListCount - 1
List1.ListIndex = i
Print #1, List1.List(i) & Chr(13)
Next
'****************************************
Close #1

MsgBox "保存完毕！", vbOKOnly + vbInformation, "提示"
End Sub

Private Sub Command2_Click()
SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3
End Sub

Private Sub Command3_Click()
SetWindowPos Me.hwnd, -2, 0, 0, 0, 0, 3
End Sub

Private Sub Command4_Click()
Dim i As Long
For i = 0 To List1.ListCount - 1
List1.Selected(i) = False
Next
Dim findstr As String
findstr = InputBox("输入搜索项（这里查找项，区分大小写）", "程序提示")
If findstr = "" Then Exit Sub
For i = 0 To List1.ListCount - 1
Label2.Caption = "所在行数：" & i + 1
List1.ListIndex = i
If InStr(List1.List(i), findstr) <> 0 Then
'此处是直接匹配项,但不是关键字查找,进程任何信息变化都会查找失败,需优化
guanbijishu = MsgBox("恭喜你，找到了项————" & List1.List(i) & "行数：" & i + 1, vbOKCancel)
If guanbijishu = 2 Then
Exit Sub
End If
List1.SetFocus
List1.Selected(i) = True
End If
Next
End Sub

Private Sub Command5_Click()
On Error Resume Next
Dim proname As String, hProcess As Long
proname = InputBox("输入进程Id，十进制数值", "程序提示") '未完成关键字匹配
If proname = "" Then Exit Sub
hProcess = OpenProcess(&H1F0FFF, False, proname)
NtTerminateProcess hProcess, 0
Command11_Click
End Sub

Private Sub Command6_Click()
On Error Resume Next
Dim getid As Long, getjubing As Long, chkok As Long
getid = InputBox("请输入线程ID(将线程ID从16进制转到10进制)", "程序提示") '未完成关键字匹配
getjubing = OpenThread(THREAD_ALL_ACCESS, 1, getid)
chkok = TerminateThread(getjubing, 0)
If chkok <> 0 Then
MsgBox "OK"
Command1_Click
Else
MsgBox "Wrong"
End If
End Sub

Private Sub Command7_Click()
Call findfromlist1("")
End Sub

'不需要计算器，但是不能随意删除，否则标号有问题，之后写入新的内容再重新使用该按钮
Private Sub Command8_Click()
Shell "calc.exe"
End Sub

Private Sub Command9_Click()
List2.Clear
Text1.Text = ""
End Sub

Private Sub Form_Load()
Dim bVersion(3) As Byte, jishuqi As Long
GetMem4 VarPtr(GetVersion), bVersion(0) '先检查了内核版本,毕竟使用是存在局限的
If bVersion(0) < 5 Then
MsgBox "内核版本不符", vbExclamation, "Wrong"
Unload Me
End
Else
Load Form1
Form1.Visible = True
SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3
End If
List1.FontSize = 10
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call logout
End Sub

Private Sub List1_Click()
Label2.Caption = "所在行数：" & List1.ListIndex + 1
End Sub

Private Sub List1_DblClick()
List2.AddItem List1.List(List1.ListIndex)
Call useshuipingscoll(List2)
End Sub

Private Sub List2_Click()
Text1.Text = List2.List(List2.ListIndex)
End Sub

Private Sub List2_DblClick()
List2.RemoveItem List2.ListIndex
Call useshuipingscoll(List2)
Text1.Text = ""
End Sub

Private Sub List2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
Call findfromlist1(List2.List(List2.ListIndex))
End If
End Sub

Private Sub Timer1_Timer() '用于即时刷新
Dim numjishu As Long
numjishu = 0
Dim bfBuffer() As Byte, BufferSize As Long, szInfo As String
Dim szImageName As String, CurPtr As Long
Dim SystemProcess As SYSTEM_PROCESS_INFORMATION
ZwQuerySystemInformation SystemProcessAndThreadInformation, ByVal 0&, 0, BufferSize
ReDim bfBuffer(BufferSize) As Byte
ZwQuerySystemInformation SystemProcessAndThreadInformation, bfBuffer(0), BufferSize, ByVal 0&
CurPtr = VarPtr(bfBuffer(0)) '当前指针指向第一个进程信息

Dim jishu As Long     '对进程进行计数
Do
jishu = jishu + 1
RtlMoveMemory SystemProcess, CurPtr, Len(SystemProcess)

 '开始累加 行信息 至进程Id
 '从此处修改代码，增加Hex数值，注意是增加，毕竟正常使用Dec
 
szInfo = szProcessId & CStr(SystemProcess.UniqueProcessId) '源代码，以下是修改，但是这是监视模块，以下代码复制到刷新部分

'szInfo = szProcessId & CStr(SystemProcess.UniqueProcessId) & "(Hex:" & DEC_to_HEX(CStr(SystemProcess.UniqueProcessId)) & ")"


'************************************************************************
'行信息 加长至进程映像名称并判断是否存在映像名
If SystemProcess.ImageName.Length > 0 Then
szImageName = Space(SystemProcess.ImageName.Length)
RtlMoveMemory ByVal StrPtr(szImageName), SystemProcess.ImageName.pszImageName, SystemProcess.ImageName.Length
szInfo = szInfo & szProcessName & Trim(szImageName)
Else
szInfo = szInfo & szProcessName & "<Null>"
End If
'************************************************************************
szInfo = szInfo & szThreadCount & CStr(SystemProcess.NumberOfThreads)
numjishu = numjishu + Val(CStr(SystemProcess.NumberOfThreads))
CurPtr = CurPtr + SystemProcess.NextEntryOffset '指针转移到下一个指针
Loop Until SystemProcess.NextEntryOffset = 0
Erase bfBuffer
Label1.Caption = "进程数：" & jishu
Label4.Caption = "线程总数：" + CStr(numjishu)
End Sub















'*******************************************************
'以下新建函数，若之后函数过多考虑写入模块
Public Function DEC_to_HEX(Dec As Long) As String '十进制转十六进制函数

     Dim a As String
     DEC_to_HEX = ""
     Do While Dec > 0
         a = CStr(Dec Mod 16)
         Select Case a
             Case "10": a = "A"
             Case "11": a = "B"
             Case "12": a = "C"
             Case "13": a = "D"
             Case "14": a = "E"
             Case "15": a = "F"
         End Select
         DEC_to_HEX = a & DEC_to_HEX
         Dec = Dec \ 16
     Loop
     
 End Function
