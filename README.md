##2810客户机登陆程序

####2810_tk.py

界面已经按会场座位排序了，在指定号座上输入名字即可签到。

	对单个客户端发送特殊作用指令：
	server1  --  选择服务器1
	server2  --  选择服务器2
	business --  恢复水牌界面
	refresh  --  更新会议资料

####socket_client.py

在客户端中添加到开机启动

windows:

	pyinstaller -F -w socket_client.py
将打包好的exe放到C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp目录下

linux：

在环境变量中配置

	python3 ~/socket_client.py