Dim wsh
set wsh=wscript.createobject("wscript.shell")
'�����Ǵ�����vpn��¼
wsh.run """C:\Program Files (x86)\Sangfor\SSL\SangforCSClient\SangforCSClient.exe"""
wscript.Sleep 3000
'������Ҫ���ģ������ҵ���xxx
wsh.SendKeys "xxx"
wsh.SendKeys "{TAB}"
wsh.SendKeys "123456"
wsh.SendKeys "{ENTER}"
wscript.Sleep 1000
wsh.SendKeys "123456"
wsh.SendKeys "{ENTER}"
wscript.Sleep 6000
'�򿪻�����䣬����õĲ��ǻ������·����ͬ���������е���
wsh.run """C:\Users\DELL\Desktop\Foxmail"""
wscript.Sleep 2000
wsh.SendKeys "{F2}"
'����VPN�������ҵ�VPN����BFS�������Ĳ����������޸�
'�����ʽΪrasdial vpn�� �û� ���룬������ʾ
wsh.run "cmd /c rasdial BFS aaa bbb"
wscript.Sleep 5000
'����ϵͳ�������ϣ���Attendance����������������޸�·������������¼�����Լ��ģ����Ҽ�ס�û���������,
'���������Ӧʱ����ܲ�Ψһ���������ʵ���ܲ��ˣ���ξ�ȥ�˰ɣ�
wsh.run """C:\Users\DELL\Desktop\Attendance"""
wscript.Sleep 2000
wsh.SendKeys "{ENTER}"
wscript.Sleep 2000
'wsh.SendKeys "{ENTER}"
msgbox "��ϲ�㣬�󹦸�ɣ���"