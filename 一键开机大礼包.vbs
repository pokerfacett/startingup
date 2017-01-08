Dim wsh
set wsh=wscript.createobject("wscript.shell")
'这里是打开内网vpn登录
wsh.run """C:\Program Files (x86)\Sangfor\SSL\SangforCSClient\SangforCSClient.exe"""
wscript.Sleep 3000
'根据需要更改，比如我的是xxx
wsh.SendKeys "xxx"
wsh.SendKeys "{TAB}"
wsh.SendKeys "123456"
wsh.SendKeys "{ENTER}"
wscript.Sleep 1000
wsh.SendKeys "123456"
wsh.SendKeys "{ENTER}"
wscript.Sleep 6000
'打开火狐邮箱，如果用的不是火狐或者路径不同，可以自行调整
wsh.run """C:\Users\DELL\Desktop\Foxmail"""
wscript.Sleep 2000
wsh.SendKeys "{F2}"
'外网VPN上网，我的VPN名叫BFS，如果你的不叫请自行修改
'命令格式为rasdial vpn名 用户 密码，如下所示
wsh.run "cmd /c rasdial BFS aaa bbb"
wscript.Sleep 5000
'考勤系统放桌面上，叫Attendance，如果不叫请自行修改路径，这里假设登录的是自己的，并且记住用户名和密码,
'考勤这边相应时间可能不唯一，所以如果实在受不了，这段就去了吧！
wsh.run """C:\Users\DELL\Desktop\Attendance"""
wscript.Sleep 2000
wsh.SendKeys "{ENTER}"
wscript.Sleep 2000
'wsh.SendKeys "{ENTER}"
msgbox "恭喜你，大功告成！！"