Attribute VB_Name = "Declares"
Public iniFileName As String
Public Debugmode, Loadactivex, UnloadactiveX, Beta As Boolean
Public registername, registercompany As String
Public frm As Form

Public Type SettingsetV
Enabled As Boolean
Value As Integer
End Type


Public Type SettingsetS
Enabled As Boolean
Value As String
End Type


'-----------------------------------����͸��-----------------------------------
Public glasseffectmode As Byte

'SetTransparentWindow(hwnd As Long, iTransparency As Integer)����˵����
'hwndΪ��Ҫ���õĴ�����
'iTransparencyΪ͸���ȣ�Ϊ0-100������0��ʾ��͸����100��ʾȫ͸��

Public Sub SetTransparentWindow(hwnd As Long, iTransparency As Byte)
    Dim rtn As Long
    Dim iTransform As Byte
    'iTransparencyת����SetLayeredWindowAttributes�ĵ�3����������͸���̶�(ȡֵ��Χ0 --255)
    iTransform = Int((100 - iTransparency) * 2.55)
    
    rtn = GetWindowLong(hwnd, GWL_EXSTYLE)      'ȡ�Ĵ���ԭ�ȵ���ʽ
    rtn = rtn Or WS_EX_LAYERED 'ʹ����������µ���ʽWS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, rtn    '���µ���ʽ��������
    SetLayeredWindowAttributes hwnd, 0, iTransform, LWA_ALPHA 'ע��:�Ѵ������óɰ�͸����ʽ , ��3������iTransform��ʾ͸���̶ȣ�ȡֵ��Χ0 --255, Ϊ0ʱ����һ��ȫ͸���Ĵ�����
    
End Sub

'ʹ�÷�����SetTransparentWindow Me.hwnd, 60 '�޸����е�60Ϊ0����͸��

