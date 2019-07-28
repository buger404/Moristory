VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TipPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   页面绘制器
    Dim Page As GPage
    Dim CtrlText As String, GoalText As String, RequireText As String, MyGame As String
'======================================

Public Sub Update()
    If Err.number <> 0 Then ErrCrash Err.number, Err.Description
    On Error Resume Next

    '游戏每一帧的过程（绘制过程请放在此处）
    
    Page.Clear argb(255, 0, 0, 0) '清空画布

    If MyGame <> "final" Then Page.DrawImage "bg.png", 0, 0, alpha:=1
    
    Page.Paint 2, GW / 2, GH / 2, 800, 500, argb(220, 255, 255, 255), 36, Pos:=posOnCenter
    'Page.Paint 0, GW / 2 - 400, GH / 2 - 250, 800, 50, argb(255, 0, 190, 169)
    
    If MyGame <> "final" Then Page.DrawImage "icon.png", GW / 2, GH / 2 - 250, alpha:=1, Pos:=posOnCenter
    
    'Page.Writes "游戏正式开始", GW / 2, GH / 2 - 250 + 12, 18, argb(255, 27, 27, 27), align:=StringAlignmentCenter
    
    Page.Writes "操作", GW / 2 - 400 + 40, GH / 2 - 250 + 90, 18, IIf(MyGame = "final", argb(255, 255, 0, 0), argb(255, 27, 27, 27)), style:=FontStyleBold
    Page.Writes CtrlText, GW / 2 - 400 + 60 + 60, GH / 2 - 250 + 90, 18, argb(170, 27, 27, 27)
    Page.Writes "目标", GW / 2 - 400 + 40, GH / 2 - 250 + 90 + 90, 18, IIf(MyGame = "final", argb(255, 255, 0, 0), argb(255, 27, 27, 27)), style:=FontStyleBold
    Page.Writes GoalText, GW / 2 - 400 + 60 + 60, GH / 2 - 250 + 90 + 90, 18, argb(170, 27, 27, 27)
    Page.Writes "规则", GW / 2 - 400 + 40, GH / 2 - 250 + 90 + 180, 18, IIf(MyGame = "final", argb(255, 255, 0, 0), argb(255, 27, 27, 27)), style:=FontStyleBold
    Page.Writes RequireText, GW / 2 - 400 + 60 + 60, GH / 2 - 250 + 90 + 180, 18, argb(170, 27, 27, 27)
    
    If Page.ShowSimpleButton("btn.png", GW / 2 - 97, GH - 220, IIf(MyGame = "final", "最后一战", "好的"), argb(255, 255, 255, 255), 18) = 3 Then
        Call GoGame
    End If
End Sub
Public Sub GoGame()
    Select Case MyGame
        Case "novel"
            ECore.NewTransform transFallDark, 1000, "NovelPage"
        Case "snowstorm"
            MazePage.LoadMap "snowstorm"
            ECore.NewTransform transFallDark, 3000, "MazePage"
        Case "snowman"
            ECore.NewTransform transFallDark, 3000, "SnowmanPage"
        Case "mazelake"
            MazePage.LoadMap "lake"
            ECore.NewTransform transFallDark, 3000, "MazePage"
        Case "mazecore"
            MazePage.LoadMap "core"
            ECore.NewTransform transFallDark, 3000, "MazePage"
        Case "mazesnow"
            MazePage.LoadMap "snow"
            ECore.NewTransform transFallDark, 3000, "MazePage"
        Case "maze"
            MazePage.LoadMap
            ECore.NewTransform transFallDark, 3000, "MazePage"
        Case "battle"
            ECore.NewTransform transFallDark, 3000, "BattlePage"
        Case "chess"
            ECore.NewTransform transFallDark, 3000, "TicTacToePage"
        Case "bxbattle"
            ECore.NewTransform transFallDark, 3000, "BXBattlePage"
        Case "fly"
            ECore.NewTransform transToDownFade, 1000, "FlyPage"
        Case "tlp"
            ECore.NewTransform transToDownFade, 1000, "TLPPage"
        Case "dance"
            ECore.NewTransform transDarkTo, 1000, "DancePage"
        Case "final"
            SE.Play "Thunder9.mp3"
            SE.Play "Thunder9.mp3"
            ECore.NewTransform 14, 3000, "FinalPage"
    End Select
End Sub
Public Sub ShowTip(Game As String)
    MyGame = Game
    Select Case MyGame
        Case "novel"
            CtrlText = "按Z或Enter，或点击窗口继续对话" & vbCrLf & "按X查看历史对话。" & vbCrLf & "按S打开游戏设置。"
            GoalText = "完成这个故事。"
            RequireText = "活下去。" & vbCrLf & _
                          "赢得角色们的一致好感~"
        Case "snowstorm"
            CtrlText = "按键盘中的 ↑ ↓ ← → 操控你自己在地图上移动。" & vbCrLf & "按 Z 攻击雪狼，按 X 防御。"
            GoalText = "在自己的血量大于300的前提下，将雪狼的血量降到300以下。"
            RequireText = "当雪狼出现在你的视野中时，你可以攻击雪狼，但不可以回血；" & vbCrLf & _
                          "当雪狼不在你的视野中时，你则不可以攻击雪狼，但可以回血；" & vbCrLf & _
                          "攻击和回血都会消耗你的灵力值（AP），长时间不攻击和回血可以补充灵力。"
        Case "snowman"
            CtrlText = "按住鼠标等待雪球增长到你需要的大小时，松开鼠标。"
            GoalText = "使雪人高度大于一定高度（云层高度）"
            RequireText = "按下鼠标后，雪球会持续增大。" & vbCrLf & _
                          "雪球大小如果比上一个雪球大小偏大，雪人将会倒塌。"
        Case "mazelake"
            CtrlText = "按键盘中的 ↑ ↓ ← → 操控你自己在地图上移动。"
            GoalText = "达到地图中用绿色方块标注的终点。"
            RequireText = "地图上将会记录你已经走过的地方。" & vbCrLf & "地图上的传送门将把你带去另一个地方。" & vbCrLf & "水流会把你冲走！"
        Case "mazecore"
            CtrlText = "按键盘中的 ↑ ↓ ← → 操控你自己在地图上移动。"
            GoalText = "达到地图中用绿色方块标注的终点。"
            RequireText = "地图上将会记录你已经走过的地方。" & vbCrLf & "地图上的传送门将把你带去另一个地方。"
        Case "mazesnow"
            CtrlText = "按键盘中的 ↑ ↓ ← → 操控你自己在地图上移动。"
            GoalText = "达到地图中用绿色方块标注的终点。"
            RequireText = "地图上将会记录你已经走过的地方。"
        Case "maze"
            CtrlText = "按键盘中的 ↑ ↓ ← → 操控你自己在地图上移动。"
            GoalText = "达到地图中用绿色方块标注的终点。"
            RequireText = "地图上将会记录你已经走过的地方。"
        Case "battle"
            CtrlText = "按窗口中的按钮攻击或回血。"
            GoalText = "在自己的血量大于300的前提下，将卡西的血量降到300以下。"
            RequireText = "按下攻击后，你将有一定的时间攻击，" & vbCrLf & _
                          "当卡西进入瞄准准心时，按下鼠标。" & vbCrLf & _
                          "击中位置越接近卡西中心，造成的伤害越大。"
        Case "chess"
            CtrlText = "在你的回合时，按下九个格子中任意一个空白位置下棋。"
            GoalText = "在兮兮达成自己下的三个棋子可以连成一条直线的目标之前，" & vbCrLf & _
                       "使你自己下的三个棋子可以连成一条直线。"
            RequireText = "下的三个棋子可以连成一条直线的一方获胜。"
        Case "bxbattle"
            CtrlText = "移动鼠标接住攻击。"
            GoalText = "抵挡住所有攻击。"
            RequireText = "你的身体只能承受三次攻击。" & vbCrLf & _
                          "能力有限，你只能接住距离你比较近的攻击，" & vbCrLf & "为了帮助你测量距离，我在可以接住的位置涂抹上了黑色矩形。" & vbCrLf & _
                          "你只能接住距离你最近的攻击。"
        Case "fly"
            CtrlText = "按鼠标或键盘上的空格使芽子往上飞。"
            GoalText = "安全通过所有障碍。"
            RequireText = "你可以撞到柱子3次而安然无恙。" & vbCrLf & _
                          "尽量不要去撞障碍物，很疼。" & vbCrLf & _
                          "别让自己摔死了。"
        Case "tlp"
            CtrlText = "选择你手中的四张牌进行攻击。"
            GoalText = "在自己的血量高于300的前提下，将塔克的血量降到300以下。"
            RequireText = "每个回合你可以抽取3~6张牌。" & vbCrLf & _
                          "每个回合你可以使用4张牌。" & vbCrLf & _
                          "塔克有时候会吓到你使你无法行动和抽牌。"
        Case "dance"
            CtrlText = "按照指示按键盘上的↑↓←→"
            GoalText = "完成这场舞蹈。"
            RequireText = "雅月将会和你一起跳舞。" & vbCrLf & "雅月将会告诉你接下来的动作。" & vbCrLf & "你只能失误5次。"
        Case "final"
            CtrlText = "按键盘上的↑↓←→操作，选取攻击和避开弹幕。"
            GoalText = "？？？？？"
            RequireText = "？？？？？？？？？？？？"
        Case "ending"
            ECore.NewTransform transFallDark, 1000, "EndingPage"
    End Select
End Sub
Public Sub Wheel(Direction As Integer, Depth As Single)
    '鼠标滚轮事件
    'Direction:方向，Depth：深度
End Sub
Public Sub AnimationMsg(id As String, msg As String)
    '动画消息接收
End Sub
Public Sub Enter()
    '页面进入事件
End Sub
Public Sub Leave()
    '页面离开事件
End Sub
Public Sub Gotfocus()
    '窗口取得焦点事件
End Sub
Public Sub Lostfocus()
    '窗口失去焦点事件
End Sub
Public Sub AnimationDone(id As Integer)
    '动画播放结束事件
    'id：结束的动画ID
End Sub
Private Sub Class_Initialize()
    '创建页面绘制器
    Set Page = New GPage
    Page.Create Me
    '导入游戏资源
    Page.Res.newImage App.Path & "\assets\bg\DarkSpace.png", GW, GH, "bg.png"
    Page.Res.NewImages App.Path & "\assets\tipui"
    '创建页面
    ECore.Add Page, "TipPage"
    
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub
