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


'-----------------------------------窗体透明-----------------------------------
Public glasseffectmode As Byte

'SetTransparentWindow(hwnd As Long, iTransparency As Integer)参数说明：
'hwnd为所要设置的窗体句柄
'iTransparency为透明度，为0-100的数，0表示不透明，100表示全透明

Public Sub SetTransparentWindow(hwnd As Long, iTransparency As Byte)
    Dim rtn As Long
    Dim iTransform As Byte
    'iTransparency转换成SetLayeredWindowAttributes的第3个参数，即透明程度(取值范围0 --255)
    iTransform = Int((100 - iTransparency) * 2.55)
    
    rtn = GetWindowLong(hwnd, GWL_EXSTYLE)      '取的窗口原先的样式
    rtn = rtn Or WS_EX_LAYERED '使窗体添加上新的样式WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, rtn    '把新的样式赋给窗体
    SetLayeredWindowAttributes hwnd, 0, iTransform, LWA_ALPHA '注释:把窗体设置成半透明样式 , 第3个参数iTransform表示透明程度，取值范围0 --255, 为0时就是一个全透明的窗体了
    
End Sub

'使用方法：SetTransparentWindow Me.hwnd, 60 '修改其中的60为0，则不透明

