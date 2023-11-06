'1 Excel有哪些命令栏
Sub 列出所有命令栏()
    Dim Index As Long
    Dim CommandBarType(1 To 3) As String
    CommandBarType(1) = "msoBarTypeNormal"
    CommandBarType(2) = "msoBarTypeMenuBar"
    CommandBarType(3) = "msoBarTypePopup"
    For Index = 1 To Application.CommandBars.Count
        With Application.CommandBars(Index)
            Cells(Index + 1, 1) = Index
            Cells(Index + 1, 2) = .Name         '英文名
            Cells(Index + 1, 3) = .NameLocal    '本地化名称
            Cells(Index + 1, 4) = CommandBarType(.Type + 1)
            Cells(Index + 1, 5) = .BuiltIn      '是否为内置工具栏
        End With
    Next Index
End Sub
'2 添加新的命令栏
'.Add(Name, Position, MenuBar, Temporary)
    'Name:命令栏的名称
    'Position：         命令栏显示的位置
        'msoBarLeft、msoBarTop、msoBarRight、msoBarBottom        
        'msoBarFloating 新命令栏不固定
        'msoBarPopup    新命令栏为快捷菜单
    'MenuBar：          布尔，是否替换活动菜单栏
    'Temporary：        是否为临时命令栏（Excel 关闭后是否会自动删除)
Sub 添加简单工具栏()
    Dim myBAR As CommandBar
    Set myBAR = Application.CommandBars.Add("我的命令栏", msoBarLeft, False, True)
    myBAR.Visible = True    '添加后要显示出来才能看到
End Sub

'3 删除命令栏
Sub 删除命令栏()
    Dim myBAR As CommandBar
    Set myBAR = Application.CommandBars("我的命令栏")
    myBAR.Delete
End Sub
'4 恢复命令栏的默认设置
Sub 恢复命令栏默认()
    Dim myBAR As CommandBar
    Set myBAR = CommandBars("我的命令栏")
    myBAR.Reset
End Sub
'5 屏蔽所有命令栏中复制命令
Sub DisableAllCopyCommand()
    ' 命令 ID 具有唯一性
    Dim combars As CommandBarControls
    Dim combar As CommandBarControl
    Dim k As Long, ID_num As Long
    ID_num = Application.CommandBars(1).Controls("编辑(&E)").Controls("复制(&C)").ID
    Set combars = Application.CommandBars.FindControls(ID:=ID_num)
    For Each combar In combars
        combar.Enabled = False
    Next combar
End Sub
'6 在命令栏中添加命令
Sub 添加命令()
    On Error Resume Next
    Dim myBAR As CommandBarButton
    Application.CommandBars("Cell").Controls("我的命令").Delete
    Set myBAR = Application.CommandBars("Cell").Controls.Add(before:=1) '添加到最上的位置
    With myBAR
        .Caption = "我的命令"
        .BeginGroup = True                  '添加分组线
        .FaceId = 199                       '显示的图标
        .Style = msoButtonIconAndCaption    '图标和文字的显示
        .OnAction = "ABC"                   '指定要运行的宏
    End With
End Sub
'7 在命令栏中添加组合框
' 点击该控件，出现列表以待选取项目
' 选择项目后，会自动运行宏
Dim mycom As CommandBarComboBox
Sub 添加组合框()
    On Error Resume Next
    Dim Index as Long
    Application.CommandBars("CELL").Controls("工作表显示").Delete
    Set mycom = Application.CommandBars("cell").Controls.Add(Type:=msoControlComboBox, before:=1) '添加到最上的位置
    With mycom
        .Caption = "工作表显示"
        .BeginGroup = True              ' 添加分组线
        .OnAction = "选取工作表"        ' 指定要运行的宏
        .Width = 100
        .DropDownWidth = 70
        .Text = Sheets(1).Name
        For Index = 1 To Sheets.Count
            .AddItem Sheets(Index).Name '添加项目
        Next Index
    End With
End Sub
Sub 选取工作表()
    Sheets(mycom.Text).Select
End Sub
'8 添加多级菜单
Sub 添加子菜单()
    On Error Resume Next
    Dim Index
    Dim ParentComPopup As CommandBarPopup
    Dim ChildPopup As CommandBarPopup
    Dim comBtn As CommandBarButton
    Application.CommandBars("CELL").Controls("工作表显示").Delete
    ' 添加父菜单，弹出式控件
    Set ParentComPopup = Application.CommandBars("cell").Controls.Add(Type:=msoControlPopup, before:=1) '添加到最上的位置
    With ParentComPopup
        .Caption = "工作表显示"
        .BeginGroup = True '添加分组线
    End With
    Set comBtn = ParentComPopup.Controls.Add(before:=1)  '添加到最上的位置
    With comBtn
        .Caption = "复制工作表"
        .FaceId = 22
        .Style = msoButtonIconAndCaption '图标和文字的显示
    End With
    Set comBtn = ParentComPopup.Controls.Add(before:=2)  '添加到最上的位置
    With comBtn
        .Caption = "删除工作表"
        .FaceId = 20
        .Style = msoButtonIconAndCaption '图标和文字的显示
    End With
    Set ChildPopup = ParentComPopup.Controls.Add(Type:=msoControlPopup, before:=3) '添加到最上的位置
    With ChildPopup
        .Caption = "移动工作表"
    End With
End Sub
Sub 删除命令()
    On Error Resume Next
    With Application.CommandBars("CELL").Controls(1)
        If .BuiltIn = False Then .Delete
    End With
End Sub
'9 添加自定义快捷菜单
Sub 添加快捷菜单()
    Dim mypup As CommandBar
    Dim com As CommandBarButton
    Dim x
    Application.CommandBars("ABC").Delete
    Set mypup = Application.CommandBars.Add(Name:="ABC", Position:=msoBarPopup)
    For x = 1 To 4
        Set com = mypup.Controls.Add
        com.Caption = Choose(x, "兰色幻想", "小妖", "小佩", "展翅")
        com.FaceId = 17 + x
        com.OnAction = "A"
    Next x
End Sub
' 通过以下命令调用快捷菜单
Application.CommandBars("ABC").ShowPopup
'10 获取内置图标
    '用 copyFace 方法复制命令按钮的图标，然后贴到其他命令按钮上或单元格
Sub FaceId()
    Application.ScreenUpdating = False
    Dim x As Integer, Y As Integer, k As Integer
    On Error Resume Next
    Dim 控件 As CommandBarButton
    Set 控件 = Application.CommandBars(4).Controls.Add
    For x = 1 To 10
        For Y = 1 To 5
            k = k + 1
            Sheets("图标").Cells(x, Y) = k
            控件.FaceId = k
            控件.CopyFace
            Sheets("图标").Cells(x, Y).Select
            ActiveSheet.Paste
        Next Y
    Next x
    控件.Delete
    Set 控件 = Nothing
    Application.ScreenUpdating = True
End Sub

'11 自定义图标

Sub 添加()
    Call 删除图片()
    Call 添加图片()
    Dim Mcom As CommandBar
    Dim Mbotton As CommandBarButton
    Dim i As Integer
    On Error Resume Next
    Application.CommandBars("工具栏").Delete
    Set Mcom = Application.CommandBars.Add
    Mcom.Visible = True
    Mcom.Name = "工具栏"
    For i = 1 To 4
        Set Mbotton = Mcom.Controls.Add
        Sheet1.Pictures(i).Copy
        Mbotton.PasteFace
    Next i
    Call 删除图片()
End Sub
Sub 插入图片()
    Dim x, bbb
    For x = 1 To 4
    Sheets("SHEET1").Pictures.Insert ThisWorkbook.Path & "\" & x & ".jpg"
    Next x
End Sub
Sub 删除图片()
    On Error Resume Next
    Dim xx
    For Each xx In Sheets("SHEET1").Pictures
        xx.Delete
    Next xx
End Sub