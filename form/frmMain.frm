VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "打砖块"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   11760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   562
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   784
   StartUpPosition =   2  '屏幕中心
   Begin VB.Menu 开始游戏 
      Caption         =   "开始游戏"
   End
   Begin VB.Menu 玩法说明 
      Caption         =   "玩法说明"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---API接口---
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Function sndPlaySoundFromMemory Lib "winmm.dll" Alias "sndPlaySoundA" (lpszSoundName As Any, ByVal uFlags As Long) As Long
Private Const SND_ASYNC As Long = &H1
Private Const SND_MEMORY As Long = &H4
'---游戏全局---
Dim mGame_State As Game_Status
Dim mPaddle_Direction As Paddle_Direction
Dim Data_Survey As Boolean
Dim Score As Single
Dim fps As Single
Dim Control_interval As Single
'---游戏音效---
Dim Paddle_Audio() As Byte
Dim Score_Coin() As Byte
Dim Loser_Audio() As Byte
'---弹珠数据---
Dim Ball_R As Single
Dim Ball_X As Single
Dim Ball_Y As Single
Dim Ball_Direction_X As Single
Dim Ball_Direction_Y As Single
Dim Ball_Speed As Single
'---方块数据---
Dim Block_X(0 To 32) As Single
Dim Block_Y(0 To 32) As Single
Dim Block_Valid(0 To 32) As Boolean
Dim Block_H As Single
Dim Block_W As Single
Dim Block_Map_Index As Long
'---挡板数据---
Dim Paddle_L As Single
Dim Paddle_X As Single
Dim Paddle_Y As Single
Dim Paddle_Unit As Single '挡板长度单元
Dim Paddle_Speed As Single
'---列表数据---
Private Enum Game_Status
    Game_STATE_RUNNING = 0
    Game_STATE_PAUSE = 1
    Game_STATE_STOP = 2
End Enum

Private Enum Paddle_Direction
    Paddle_Left = 1
    Paddle_Right = 2
    Paddle_Stop = 0
End Enum

'---实现代码---
Private Sub Form_Initialize()
    Me.Height = 600 * Screen.TwipsPerPixelY
    Me.Width = 800 * Screen.TwipsPerPixelX
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyA Or KeyCode = vbKeyLeft Then
        mPaddle_Direction = Paddle_Left
    ElseIf KeyCode = vbKeyD Or KeyCode = vbKeyRight Then
        mPaddle_Direction = Paddle_Right
    ElseIf KeyCode = vbKeyS Or KeyCode = vbKeyDown Then
        mPaddle_Direction = Paddle_Stop
    End If
    If KeyCode = vbKeyF And Shift = 2 Then
        Data_Survey = Not (Data_Survey)
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    mPaddle_Direction = Paddle_Stop
End Sub

Private Sub Game_Load()
    Score = 0
    Paddle_X = Paddle_L * 2
    Paddle_Y = 480
    Ball_X = Paddle_X + Paddle_L / 2
    Ball_Y = Paddle_Y - Ball_R
    Ball_Direction_X = Ball_Speed
    Ball_Direction_Y = -Ball_Speed
End Sub


Private Sub Form_Load() '数据初始化
    '控制数据
    mGame_State = Game_STATE_STOP
    Control_interval = 80
    fps = 80
    Data_Survey = False
    Score = 0
    Paddle_Unit = 5
    '挡板数据
    Paddle_L = 30 * Paddle_Unit
    Paddle_X = Paddle_L * 2
    Paddle_Y = 480
    Paddle_Speed = Paddle_Unit * 2
    '球数据
    Ball_R = 8
    Ball_X = Paddle_X + Paddle_L / 2
    Ball_Y = Paddle_Y - Ball_R
    Ball_Speed = Paddle_Unit
    Ball_Direction_X = Ball_Speed
    Ball_Direction_Y = -Ball_Speed
    '方块
    Block_W = 50
    Block_H = 20
    Block_Map_Index = 1
    '音效
    Paddle_Audio = LoadResData("PADDLE.WAV", "AUDIO")
    Score_Coin = LoadResData("COIN.WAV", "AUDIO")
    Loser_Audio = LoadResData("LOSER.WAV", "AUDIO")
End Sub

Private Sub 开始游戏_Click()
    If 开始游戏.Caption = "开始游戏" Or 开始游戏.Caption = "继续游戏" Then
        开始游戏.Caption = "暂停游戏"
        mGame_State = Game_STATE_RUNNING
        Game_Block_Map Block_Map_Index
        Call Game_Loop
    ElseIf 开始游戏.Caption = "暂停游戏" Then
        开始游戏.Caption = "继续游戏"
        mGame_State = Game_STATE_PAUSE
    End If
End Sub

Private Sub Game_Loop() '进入游戏循环
    Dim fps_Time_New As Long
    Dim fps_Time_Last As Long
    Dim Control_Time_New As Long
    Dim Control_Time_Last As Long
    
    Dim i As Integer
    Dim n As Integer
    While DoEvents
        If mGame_State = Game_STATE_RUNNING Then
            fps_Time_New = timeGetTime()
            '控制实现
            If fps_Time_New - fps_Time_Last >= 1000 / fps Then
                fps_Time_Last = fps_Time_New
                'UI实现
                Me.Cls
                Call Frame_Draw
                Call Game_Draw
            End If
            Control_Time_New = timeGetTime()
            If Control_Time_New - Control_Time_Last >= 1000 / Control_interval Then
                Control_Time_Last = Control_Time_New
                '挡板左右控制
                If mPaddle_Direction = Paddle_Left Then
                    If Paddle_X >= Paddle_Speed Then Paddle_X = Paddle_X - Paddle_Speed
                ElseIf mPaddle_Direction = Paddle_Right Then
                    If Paddle_X < 599 - Paddle_L Then Paddle_X = Paddle_X + Paddle_Speed
                End If
                '球运行
                If Ball_X <= Ball_R Then Ball_Direction_X = -Ball_Direction_X
                If Ball_Y <= Ball_R Then Ball_Direction_Y = -Ball_Direction_Y
                If Ball_X >= 599 - Ball_R Then Ball_Direction_X = -Ball_Direction_X
                '球与挡板碰撞
                If Ball_Y > Paddle_Y - Ball_R And Ball_X >= Paddle_X And Ball_X <= Paddle_X + Paddle_L And Ball_Y <= Paddle_Y + Ball_R Then
                    '播放音效
                    sndPlaySoundFromMemory Paddle_Audio(0), SND_ASYNC Or SND_MEMORY
                    Ball_Direction_Y = -Ball_Direction_Y
                End If
                '方块碰撞
                For i = 0 To 32
                    If Block_Valid(i) = True Then
                        n = Collide(Block_X(i), Block_Y(i), Block_W, Block_H, Ball_X, Ball_Y, Ball_R)
                        If n = 1 Or n = 2 Then
                            Ball_Direction_X = -Ball_Direction_X
                            Score = Score + 1
                            sndPlaySoundFromMemory Score_Coin(0), SND_ASYNC Or SND_MEMORY
                            Block_Valid(i) = False
                            Exit For
                        ElseIf n = 3 Or n = 4 Then
                            Ball_Direction_Y = -Ball_Direction_Y
                            Score = Score + 1
                            sndPlaySoundFromMemory Score_Coin(0), SND_ASYNC Or SND_MEMORY
                            Block_Valid(i) = False
                            Exit For
                        End If
                    End If
                Next i
                Ball_X = Ball_X + Ball_Direction_X
                Ball_Y = Ball_Y + Ball_Direction_Y
                '胜利判断
                n = 0
                For i = 0 To 32
                    If Block_Valid(i) = True Then
                        n = n + 1
                    End If
                Next i
                If n = 0 Then
                    Me.FontSize = 16
                    Me.ForeColor = vbRed
                    Me.CurrentX = 625
                    Me.CurrentY = 41
                    Me.Print "游戏胜利"
                    mGame_State = Game_STATE_PAUSE
                End If
                '失败判断
                If Ball_Y - Ball_R >= 600 Then
                    Me.FontSize = 16
                    Me.ForeColor = vbRed
                    Me.CurrentX = 625
                    Me.CurrentY = 41
                    Me.Print "游戏结束"
                    开始游戏.Caption = "开始游戏"
                    sndPlaySoundFromMemory Loser_Audio(0), SND_ASYNC Or SND_MEMORY
                    mGame_State = Game_STATE_STOP
                    '清空数据
                    Call Game_Load
                    Exit Sub
                End If
            End If
        End If
        Sleep (2)
    Wend
End Sub
'矩形圆形碰撞检测
Private Function Collide(ByVal X1 As Single, ByVal Y1 As Single, ByVal W As Single, ByVal H As Single, ByVal X2 As Single, ByVal Y2 As Single, ByVal R As Single) As Long
    Dim mX As Single
    Dim mY As Single
    Dim L As Single
    'Collide 0 无碰撞 1 左 2 右 3 上 4 下
    If X2 < X1 Then
        mX = X1
        Collide = 1
    ElseIf X2 > X1 + W Then
        mX = X1 + W
        Collide = 2
    Else
        mX = X2
    End If
    If Y2 < Y1 Then
        mY = Y1
        Collide = 3
    ElseIf Y2 > Y1 + H Then
        mY = Y1 + H
        Collide = 4
    Else
        mY = Y2
    End If
    L = (mX - X2) ^ 2 + (mY - Y2) ^ 2
    If L > R ^ 2 Then
        Collide = 0
    End If
End Function

Private Sub Frame_Draw()
    Me.ForeColor = vbBlack
    Me.DrawWidth = 1
    Me.Line (600, 0)-(600, 600)
    Me.FontSize = 16
    Me.CurrentX = 620
    Me.CurrentY = 20
    Me.Print "分数：" & Score
    '监测数据
    If Data_Survey = True Then
        Me.FontSize = 10
        Me.CurrentX = 620
        Me.CurrentY = 180
        Me.Print "数据监测"
        Me.CurrentX = 620
        Me.CurrentY = 200
        Me.Print "mPaddle_Direction:" & mPaddle_Direction
        Me.CurrentX = 620
        Me.CurrentY = 220
        Me.Print "Paddle_X:" & Paddle_X
        Me.CurrentX = 620
        Me.CurrentY = 240
        Me.Print "Ball_Direction_X:" & Ball_Direction_X
        Me.CurrentX = 620
        Me.CurrentY = 260
        Me.Print "Ball_X:" & Ball_X
    End If
    Me.FontSize = 10
    Me.FontName = "微软雅黑"
    Me.CurrentX = 620
    Me.CurrentY = 500
    Me.Print "作者：0yufan0@VB吧"
    Me.CurrentX = 620
    Me.CurrentY = 520
    Me.Print "邮箱：woyufan@163.com"
End Sub

Private Sub Game_Block_Map(ByVal Map_Index As Long)  '地图生成
    Dim i As Integer
    Dim j As Integer
    Select Case Map_Index
        Case 1
            For i = 0 To 7
                Block_X(i) = 100 + Block_W * i
                Block_Y(i) = 60
                Block_Valid(i) = True
            Next i
            For i = 8 To 13
                Block_X(i) = 100 + Block_W * (i - 7)
                Block_Y(i) = 60 + Block_H
                Block_Valid(i) = True
            Next i
            For i = 14 To 17
                Block_X(i) = 100 + Block_W * (i - 12)
                Block_Y(i) = 60 + 2 * Block_H
                Block_Valid(i) = True
            Next i
            For i = 18 To 19
                Block_X(i) = 100 + Block_W * (i - 15)
                Block_Y(i) = 60 + 3 * Block_H
                Block_Valid(i) = True
            Next i
    End Select
End Sub

Private Sub Game_Draw()
    Dim i As Integer
    '画挡板
    Me.DrawWidth = 2
    Me.ForeColor = vbRed
    Me.FillStyle = 0
    Me.Line (Paddle_X, Paddle_Y)-(Paddle_X + Paddle_L, Paddle_Y)
    '画球
    Me.ForeColor = vbRed
    Me.FillColor = vbRed
    Me.Circle (Ball_X, Ball_Y), Ball_R
    '画方块
    Me.ForeColor = vbBlack
    Me.FillColor = vbBlack
    For i = 0 To 32
        If Block_Valid(i) = True Then
            Me.Line (Block_X(i), Block_Y(i))-(Block_X(i) + Block_W, Block_Y(i) + Block_H), , BF
        End If
    Next i
End Sub

Private Sub 玩法说明_Click()
    Me.Cls
    Call Frame_Draw
    Me.FontSize = 14
    Me.ForeColor = vbBlack
    Me.FillColor = vbBlack
    Me.CurrentX = 20
    Me.CurrentY = 20
    Me.Print "游戏玩法说明："
    Me.CurrentX = 20
    Me.CurrentY = 40
    Me.Print "1.通过A/D来控制挡板的移动，保证弹珠球不掉落。"
    Me.CurrentX = 20
    Me.CurrentY = 60
    Me.Print "2.当挡板移动时，弹珠球的速度会改变，甚至有时候会突然间加快。"
    Me.CurrentX = 20
    Me.CurrentY = 80
    Me.Print "3.尽量通过消除砖块获得更高的分数。"
    开始游戏.Caption = "继续游戏"
    mGame_State = Game_STATE_PAUSE
End Sub
