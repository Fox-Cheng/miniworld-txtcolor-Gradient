VERSION 5.00
Begin VB.Form 主窗体 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "迷你世界文本渐变工具"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7995
   LinkTopic       =   "主窗体"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   7995
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame 杂项设置 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "杂项设置"
      ForeColor       =   &H000000FF&
      Height          =   3015
      Left            =   100
      TabIndex        =   2
      Top             =   1200
      Width           =   7815
      Begin VB.OptionButton Color3 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "粉黄青"
         ForeColor       =   &H00FFFFFF&
         Height          =   400
         Left            =   2600
         TabIndex        =   9
         Top             =   600
         Width           =   1000
      End
      Begin VB.OptionButton Color2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "黄青粉"
         ForeColor       =   &H00FFFFFF&
         Height          =   400
         Left            =   1400
         TabIndex        =   8
         Top             =   600
         Width           =   1000
      End
      Begin VB.OptionButton Color1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "青粉黄"
         ForeColor       =   &H00FFFFFF&
         Height          =   400
         Left            =   200
         TabIndex        =   7
         Top             =   600
         Value           =   -1  'True
         Width           =   1000
      End
      Begin VB.CheckBox 超小字 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "超小字"
         ForeColor       =   &H008080FF&
         Height          =   300
         Left            =   2800
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   6
         Top             =   200
         Width           =   1000
      End
      Begin VB.CheckBox 下划线 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "下划线"
         ForeColor       =   &H008080FF&
         Height          =   300
         Left            =   1800
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   5
         Top             =   200
         Width           =   1000
      End
      Begin VB.CheckBox 倾斜 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "倾斜"
         ForeColor       =   &H008080FF&
         Height          =   300
         Left            =   1000
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   4
         Top             =   200
         Width           =   800
      End
      Begin VB.CheckBox 加粗 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "加粗"
         ForeColor       =   &H008080FF&
         Height          =   300
         Left            =   200
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   3
         Top             =   200
         Width           =   800
      End
   End
   Begin VB.TextBox 生成内容 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1000
      Left            =   4050
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "主窗体.frx":0000
      Top             =   100
      Width           =   3850
   End
   Begin VB.TextBox 输入文本 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   1000
      Left            =   100
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "主窗体.frx":0027
      Top             =   100
      Width           =   3850
   End
End
Attribute VB_Name = "主窗体"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'在选择颜色后 刷新一下生成后的内容
Private Sub Color1_Click()
    刷新
End Sub

'在选择颜色后 刷新一下生成后的内容
Private Sub Color2_Click()
    刷新
End Sub

'在选择颜色后 刷新一下生成后的内容
Private Sub Color3_Click()
    刷新
End Sub

'在超小字点击后 刷新一下生成后的内容
Private Sub 超小字_Click()
    刷新
End Sub

'在下划线点击后 刷新一下生成后的内容
Private Sub 下划线_Click()
    刷新
End Sub

'在加粗点击后 刷新一下生成后的内容
Private Sub 加粗_Click()
    刷新
End Sub

'在倾斜点击后 刷新一下生成后的内容
Private Sub 倾斜_Click()
    刷新
End Sub

'点击输入文本时,自动清理默认文本
Private Sub 输入文本_click()
    If Not Me.输入文本 Is Nothing Then
        If Mid(Me.输入文本.text, 1, 7) = "在此处输入文本" Then
            Me.输入文本.text = ""
        End If
    End If
End Sub

'输入文本实时刷新
Private Sub 输入文本_Change()
    Dim 转换后内容 As String
    转换后内容 = 添加颜色(Me.输入文本.text)
    设置粘贴板 转换后内容
    Me.生成内容.text = 转换后内容
End Sub

'安全设置粘贴板,无报错
Private Sub 设置粘贴板(ByVal text As String)
    On Error Resume Next
        Clipboard.Clear
        Clipboard.SetText text
    On Error GoTo 0
End Sub

'更新颜色,主函数
Private Function 添加颜色(内容 As String)
    
    Dim 颜色组(99, 200) As String
    Dim I As String
    Dim B As String
    Dim U As String
    Dim SIZE As String
    Dim COLOR As Integer
    Dim 返回值 As String
    Dim 字符长度 As Integer
    Dim 内容长度 As Integer
    Dim 分隔长度 As Integer
    Dim 转换中 As String
    Dim A As Integer
    
    '颜色-青粉黄
    颜色组(0, 0) = "#cAAFFFF"
    颜色组(0, 1) = "#cBBEEFF"
    颜色组(0, 2) = "#cDDDDFF"
    颜色组(0, 3) = "#cEEBBFF"
    颜色组(0, 4) = "#cFFAAFF"
    颜色组(0, 5) = "#cFFBBEE"
    颜色组(0, 6) = "#cFFDDDD"
    颜色组(0, 7) = "#cFFEEBB"
    颜色组(0, 8) = "#cFFFFAA"
    
    '颜色-黄青粉
    颜色组(1, 0) = "#cFFFFAA"
    颜色组(1, 1) = "#cEEFFBB"
    颜色组(1, 2) = "#cDDFFDD"
    颜色组(1, 3) = "#cBBFFEE"
    颜色组(1, 4) = "#cAAFFFF"
    颜色组(1, 5) = "#cBBEEFF"
    颜色组(1, 6) = "#cDDDDFF"
    颜色组(1, 7) = "#cEEBBFF"
    颜色组(1, 8) = "#cFFAAFF"
    
    '颜色-粉黄青
    颜色组(2, 0) = "#cFFAAFF"
    颜色组(2, 1) = "#cFFBBEE"
    颜色组(2, 2) = "#cFFDDDD"
    颜色组(2, 3) = "#cFFEEBB"
    颜色组(2, 4) = "#cFFFFAA"
    颜色组(2, 5) = "#cEEFFBB"
    颜色组(2, 6) = "#cDDFFDD"
    颜色组(2, 7) = "#cBBFFEE"
    颜色组(2, 8) = "#cAAFFFF"
    '创建定义
    
    字符长度 = 0
    If Me.倾斜.Value Then
        I = "[I]"
        字符长度 = 字符长度 + 3
    Else
        I = ""
    End If
    
    If Me.下划线.Value Then
        U = "[U]"
        字符长度 = 字符长度 + 3
    Else
        U = ""
    End If
    
    If Me.加粗.Value Then
        B = "[B]"
        字符长度 = 字符长度 + 3
    Else
        B = ""
    End If
    
    If Me.超小字.Value Then
        SIZE = "[SIZE]"
        字符长度 = 字符长度 + 6
    Else
        SIZE = ""
    End If

    If Me.Color1 Then
        COLOR = 0
    ElseIf Me.Color2 Then
        COLOR = 1
    ElseIf Me.Color3 Then
        COLOR = 2
    End If
    '获取杂项设置
    
    分隔长度 = 1
    内容长度 = Len(内容) * 3 + (Int(Len(内容) / 分隔长度) * 8)
    While 内容长度 + 字符长度 > 120 And 分隔长度 < 10
        分隔长度 = 分隔长度 + 1
        内容长度 = Len(内容) * 3 + (Int(Len(内容) / 分隔长度) * 8)
    Wend
    
    转换中 = I & U & B & SIZE
    For A = 0 To 向上取整(Len(内容) / 分隔长度) - 1
        转换中 = 转换中 & 颜色组(COLOR, A) & Mid(内容, 1 + (A * 分隔长度), 分隔长度)
    Next A
    
    If 分隔长度 > 9 Then
        返回值 = "长度超限"
        Me.生成内容.Font.SIZE = 40
        Me.生成内容.ForeColor = vbRed
    Else
        返回值 = 转换中
        Me.生成内容.Font.SIZE = 10
        Me.生成内容.ForeColor = vbWhite
    End If
    '构造渐变字
    
    添加颜色 = 返回值
    '返回
End Function


Function 刷新()
    Dim 转换后内容 As String
    转换后内容 = 添加颜色(Me.输入文本.text)
    设置粘贴板 转换后内容
    Me.生成内容.text = 转换后内容
End Function


Function 向上取整(ByVal num As Double) As Double
    If num = Fix(num) Then
        ' 本身是整数，直接返回
        向上取整 = num
    Else
        ' 有小数部分，正数+1，负数保持整数部分
        向上取整 = IIf(num > 0, Fix(num) + 1, Fix(num))
    End If
End Function
