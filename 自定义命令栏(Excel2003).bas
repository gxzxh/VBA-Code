'1 Excel����Щ������
Sub �г�����������()
    Dim Index As Long
    Dim CommandBarType(1 To 3) As String
    CommandBarType(1) = "msoBarTypeNormal"
    CommandBarType(2) = "msoBarTypeMenuBar"
    CommandBarType(3) = "msoBarTypePopup"
    For Index = 1 To Application.CommandBars.Count
        With Application.CommandBars(Index)
            Cells(Index + 1, 1) = Index
            Cells(Index + 1, 2) = .Name         'Ӣ����
            Cells(Index + 1, 3) = .NameLocal    '���ػ�����
            Cells(Index + 1, 4) = CommandBarType(.Type + 1)
            Cells(Index + 1, 5) = .BuiltIn      '�Ƿ�Ϊ���ù�����
        End With
    Next Index
End Sub
'2 ����µ�������
'.Add(Name, Position, MenuBar, Temporary)
    'Name:������������
    'Position��         ��������ʾ��λ��
        'msoBarLeft��msoBarTop��msoBarRight��msoBarBottom        
        'msoBarFloating �����������̶�
        'msoBarPopup    ��������Ϊ��ݲ˵�
    'MenuBar��          �������Ƿ��滻��˵���
    'Temporary��        �Ƿ�Ϊ��ʱ��������Excel �رպ��Ƿ���Զ�ɾ��)
Sub ��Ӽ򵥹�����()
    Dim myBAR As CommandBar
    Set myBAR = Application.CommandBars.Add("�ҵ�������", msoBarLeft, False, True)
    myBAR.Visible = True    '��Ӻ�Ҫ��ʾ�������ܿ���
End Sub

'3 ɾ��������
Sub ɾ��������()
    Dim myBAR As CommandBar
    Set myBAR = Application.CommandBars("�ҵ�������")
    myBAR.Delete
End Sub
'4 �ָ���������Ĭ������
Sub �ָ�������Ĭ��()
    Dim myBAR As CommandBar
    Set myBAR = CommandBars("�ҵ�������")
    myBAR.Reset
End Sub
'5 ���������������и�������
Sub DisableAllCopyCommand()
    ' ���� ID ����Ψһ��
    Dim combars As CommandBarControls
    Dim combar As CommandBarControl
    Dim k As Long, ID_num As Long
    ID_num = Application.CommandBars(1).Controls("�༭(&E)").Controls("����(&C)").ID
    Set combars = Application.CommandBars.FindControls(ID:=ID_num)
    For Each combar In combars
        combar.Enabled = False
    Next combar
End Sub
'6 �����������������
Sub �������()
    On Error Resume Next
    Dim myBAR As CommandBarButton
    Application.CommandBars("Cell").Controls("�ҵ�����").Delete
    Set myBAR = Application.CommandBars("Cell").Controls.Add(before:=1) '��ӵ����ϵ�λ��
    With myBAR
        .Caption = "�ҵ�����"
        .BeginGroup = True                  '��ӷ�����
        .FaceId = 199                       '��ʾ��ͼ��
        .Style = msoButtonIconAndCaption    'ͼ������ֵ���ʾ
        .OnAction = "ABC"                   'ָ��Ҫ���еĺ�
    End With
End Sub
'7 ���������������Ͽ�
' ����ÿؼ��������б��Դ�ѡȡ��Ŀ
' ѡ����Ŀ�󣬻��Զ����к�
Dim mycom As CommandBarComboBox
Sub �����Ͽ�()
    On Error Resume Next
    Dim Index as Long
    Application.CommandBars("CELL").Controls("��������ʾ").Delete
    Set mycom = Application.CommandBars("cell").Controls.Add(Type:=msoControlComboBox, before:=1) '��ӵ����ϵ�λ��
    With mycom
        .Caption = "��������ʾ"
        .BeginGroup = True              ' ��ӷ�����
        .OnAction = "ѡȡ������"        ' ָ��Ҫ���еĺ�
        .Width = 100
        .DropDownWidth = 70
        .Text = Sheets(1).Name
        For Index = 1 To Sheets.Count
            .AddItem Sheets(Index).Name '�����Ŀ
        Next Index
    End With
End Sub
Sub ѡȡ������()
    Sheets(mycom.Text).Select
End Sub
'8 ��Ӷ༶�˵�
Sub ����Ӳ˵�()
    On Error Resume Next
    Dim Index
    Dim ParentComPopup As CommandBarPopup
    Dim ChildPopup As CommandBarPopup
    Dim comBtn As CommandBarButton
    Application.CommandBars("CELL").Controls("��������ʾ").Delete
    ' ��Ӹ��˵�������ʽ�ؼ�
    Set ParentComPopup = Application.CommandBars("cell").Controls.Add(Type:=msoControlPopup, before:=1) '��ӵ����ϵ�λ��
    With ParentComPopup
        .Caption = "��������ʾ"
        .BeginGroup = True '��ӷ�����
    End With
    Set comBtn = ParentComPopup.Controls.Add(before:=1)  '��ӵ����ϵ�λ��
    With comBtn
        .Caption = "���ƹ�����"
        .FaceId = 22
        .Style = msoButtonIconAndCaption 'ͼ������ֵ���ʾ
    End With
    Set comBtn = ParentComPopup.Controls.Add(before:=2)  '��ӵ����ϵ�λ��
    With comBtn
        .Caption = "ɾ��������"
        .FaceId = 20
        .Style = msoButtonIconAndCaption 'ͼ������ֵ���ʾ
    End With
    Set ChildPopup = ParentComPopup.Controls.Add(Type:=msoControlPopup, before:=3) '��ӵ����ϵ�λ��
    With ChildPopup
        .Caption = "�ƶ�������"
    End With
End Sub
Sub ɾ������()
    On Error Resume Next
    With Application.CommandBars("CELL").Controls(1)
        If .BuiltIn = False Then .Delete
    End With
End Sub
'9 ����Զ����ݲ˵�
Sub ��ӿ�ݲ˵�()
    Dim mypup As CommandBar
    Dim com As CommandBarButton
    Dim x
    Application.CommandBars("ABC").Delete
    Set mypup = Application.CommandBars.Add(Name:="ABC", Position:=msoBarPopup)
    For x = 1 To 4
        Set com = mypup.Controls.Add
        com.Caption = Choose(x, "��ɫ����", "С��", "С��", "չ��")
        com.FaceId = 17 + x
        com.OnAction = "A"
    Next x
End Sub
' ͨ������������ÿ�ݲ˵�
Application.CommandBars("ABC").ShowPopup
'10 ��ȡ����ͼ��
    '�� copyFace �����������ť��ͼ�꣬Ȼ�������������ť�ϻ�Ԫ��
Sub FaceId()
    Application.ScreenUpdating = False
    Dim x As Integer, Y As Integer, k As Integer
    On Error Resume Next
    Dim �ؼ� As CommandBarButton
    Set �ؼ� = Application.CommandBars(4).Controls.Add
    For x = 1 To 10
        For Y = 1 To 5
            k = k + 1
            Sheets("ͼ��").Cells(x, Y) = k
            �ؼ�.FaceId = k
            �ؼ�.CopyFace
            Sheets("ͼ��").Cells(x, Y).Select
            ActiveSheet.Paste
        Next Y
    Next x
    �ؼ�.Delete
    Set �ؼ� = Nothing
    Application.ScreenUpdating = True
End Sub

'11 �Զ���ͼ��

Sub ���()
    Call ɾ��ͼƬ()
    Call ���ͼƬ()
    Dim Mcom As CommandBar
    Dim Mbotton As CommandBarButton
    Dim i As Integer
    On Error Resume Next
    Application.CommandBars("������").Delete
    Set Mcom = Application.CommandBars.Add
    Mcom.Visible = True
    Mcom.Name = "������"
    For i = 1 To 4
        Set Mbotton = Mcom.Controls.Add
        Sheet1.Pictures(i).Copy
        Mbotton.PasteFace
    Next i
    Call ɾ��ͼƬ()
End Sub
Sub ����ͼƬ()
    Dim x, bbb
    For x = 1 To 4
    Sheets("SHEET1").Pictures.Insert ThisWorkbook.Path & "\" & x & ".jpg"
    Next x
End Sub
Sub ɾ��ͼƬ()
    On Error Resume Next
    Dim xx
    For Each xx In Sheets("SHEET1").Pictures
        xx.Delete
    Next xx
End Sub