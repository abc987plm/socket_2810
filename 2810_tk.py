
__author__ = 'jun-x'

"""v1.1 update choose server and restart conference paper
v2.0 update output name list
v2.1 repair output name list to .xls
exit code after Sep.1st"""

import tkinter as tk
from socket import *
from threading import Thread
import openpyxl, datetime, tkinter.messagebox, os

# 创建窗口
window = tk.Tk()
window.title('2810客户机登陆程序')
window.geometry('500x700')

# 客户端IP顺序
ip_add = ['192.168.24.1','192.168.24.2','192.168.24.3','192.168.24.4','192.168.24.5','192.168.24.6','192.168.24.7',
            '192.168.24.8','192.168.24.9','192.168.24.10','192.168.24.11','192.168.24.12','192.168.24.13','192.168.24.14',
            '192.168.24.15','192.168.24.16','192.168.24.17','192.168.24.18','192.168.24.19','192.168.24.20','192.168.24.21',
            '192.168.24.22','192.168.24.23','192.168.24.24',]

sort = [('12','14','10','16','8','18','6','20','4','22','2'),
        ('13','11','15','9','17','7','19','5','21','3','23','1','24')]


# 发送到客户机功能
def send(ip,name):
    if name == '' or name == ' ':
        return 0
    # dest = (ip,9999)
    dest = ('127.0.0.1',9999)
    print('向%s客户机发送了%s名单' %(ip,name))
    s = socket(AF_INET,SOCK_DGRAM)
    #设置能够发送广播
    s.setsockopt(SOL_SOCKET,SO_BROADCAST,1)
    s.sendto(name.encode(),dest)
    s.close()


def outputname():
    desk = os.path.join(os.path.expanduser('~'), 'Desktop') + '\\'
    with open(desk + '参会名单.xls', 'w') as namenode:
        namenode.write('参会名单'+'\n')
        for i in all_client.values():
            if i.get() == '':
                continue
            # print('写入%s' %i.get())
            namenode.write(i.get())
            namenode.write('\n')

# 名单导入函数
def namefile():
    wb = openpyxl.load_workbook('排座顺序表.xlsx')
    ws = wb.active
    cols = []
    for col in ws.iter_cols():
        cols.append(col)
    # 交流会排法导入
    if len(cols) == 2:
        # excel表格必须只有2列，遍历A列和B列
        for i in range(2):
            try:
                # 遍历列中的元素，如果格中没有内容就路过
                for number in range(len(cols[i])):
                    if cols[i][number].value == None:
                        continue
                    # all_client客户机对应该的字典位置，在end插入cols[列][行]的值
                    all_client[sort[i][number]].insert('end',cols[i][number].value)
            except IndexError as e:
                tk.Label(window,text='第' + str(i + 1) + '排人数超额!!!',fg='red').place(x=0,y=650+i*25)

    # 主席位排法导入
    if len(cols) == 1:
        try:
            i = -1
            for numClient in all_client.values():
                i += 1
                if cols[0][i].value ==None:
                    continue
                # print(cols[0][i].value)
                numClient.insert('end',cols[0][i].value)
        except:
            tk.Label(window, text='导入的名单格式有问题，请检查!!!', fg='red').place(x=0, y=650)
    wb.close()

# 选择服务器功能
def choice_server(server):
    for i in range(len(ip_add)):
        sign = Thread(target=send,args=(ip_add[i],server))
        sign.start()
        sign.join()
# 客户机登陆测试函数
def prinf_test(a):
    pass

# 一键登陆功能
def sign_all():
    v = 0
    for i in all_client.values():
        sign = Thread(target=send,args=(ip_add[v],i.get()))
        sign.start()
        sign.join()
        v += 1

# 清除座位名单内容
def clean():
    for i in all_client.values():
        i.delete(0,'end')


# 菜单功能
menubar = tk.Menu(window)
filemenu = tk.Menu(menubar,tearoff=0)
menubar.add_cascade(label='菜单',menu=filemenu)
filemenu.add_cascade(label='导入名单',command=namefile)
filemenu.add_cascade(label='导出名单',command=outputname)
submenu = tk.Menu(filemenu,tearoff=0)
filemenu.add_cascade(label='服务器选择',menu=submenu)
submenu.add_command(label='服务器1',command=lambda :choice_server('server1'))
submenu.add_command(label='服务器2',command=lambda :choice_server('server2'))
openeation = tk.Menu(menubar,tearoff=0)
menubar.add_cascade(label='操作',menu=openeation)
openeation.add_cascade(label='刷新资料',command=lambda :choice_server('refresh'))
openeation.add_cascade(label='恢复水牌页面',command=lambda :choice_server('business'))
filemenu.add_separator()
filemenu.add_command(label='退出',command=window.quit)
window.config(menu=menubar)

# 抬头标签
tk.Label(text='投影幕布',bg='green',font=('微软雅黑',11,'bold'),width=50,).place(anchor='nw')

# 24个客户机窗口功能
z1 = tk.Entry(window,width=10,font=('微软雅黑',16),show = None)
z1.place(x=180,y=30)
z1_send = tk.Button(window,text='1号登陆',width=10,command=lambda :send(ip_add[0],z1.get()))
z1_send.place(x=280,y=30)

z2 = tk.Entry(window,width=10,font=('微软雅黑',16),show = None)
z2.place(x=270,y=70)
z2_send = tk.Button(window,text='2号登陆',width=10,command=lambda :send(ip_add[1],z2.get()))
z2_send.place(x=370,y=70)

z3 = tk.Entry(window,width=10,font=('微软雅黑',16),show = None)
z3.place(x=20,y=70)
z3_send = tk.Button(window,text='3号登陆',width=10,command=lambda :send(ip_add[2],z3.get()))
z3_send.place(x=120,y=70)

z4 = tk.Entry(window,width=10,font=('微软雅黑',16),show = None)
z4.place(x=270,y=110)
z4_send = tk.Button(window,text='4号登陆',width=10,command=lambda :send(ip_add[3],z4.get()))
z4_send.place(x=370,y=110)

z5 = tk.Entry(window,width=10,font=('微软雅黑',16),show = None)
z5.place(x=20,y=110)
z5_send = tk.Button(window,text='5号登陆',width=10,command=lambda :send(ip_add[4],z5.get()))
z5_send.place(x=120,y=110)

z6 = tk.Entry(window,width=10,font=('微软雅黑',16),show = None)
z6.place(x=270,y=150)
z6_send = tk.Button(window,text='6号登陆',width=10,command=lambda :send(ip_add[5],z6.get()))
z6_send.place(x=370,y=150)

z7 = tk.Entry(window,width=10,font=('微软雅黑',16),show = None)
z7.place(x=20,y=150)
z7_send = tk.Button(window,text='7号登陆',width=10,command=lambda :send(ip_add[6],z7.get()))
z7_send.place(x=120,y=150)

z8 = tk.Entry(window,width=10,font=('微软雅黑',16),show = None)
z8.place(x=270,y=190)
z8_send = tk.Button(window,text='8号登陆',width=10,command=lambda :send(ip_add[7],z8.get()))
z8_send.place(x=370,y=190)

z9 = tk.Entry(window,width=10,font=('微软雅黑',16),show = None)
z9.place(x=20,y=190)
z9_send = tk.Button(window,text='9号登陆',width=10,command=lambda :send(ip_add[8],z9.get()))
z9_send.place(x=120,y=190)

z10 = tk.Entry(window,width=10,font=('微软雅黑',16),show = None)
z10.place(x=270,y=230)
z10_send = tk.Button(window,text='10号登陆',width=10,command=lambda :send(ip_add[9],z10.get()))
z10_send.place(x=370,y=230)

z11 = tk.Entry(window,width=10,font=('微软雅黑',16),show = None)
z11.place(x=20,y=230)
z11_send = tk.Button(window,text='11号登陆',width=10,command=lambda :send(ip_add[10],z11.get()))
z11_send.place(x=120,y=230)

z12 = tk.Entry(window,width=10,font=('微软雅黑',16),show = None)
z12.place(x=270,y=270)
z12_send = tk.Button(window,text='12号登陆',fg='red',width=10,command=lambda :send(ip_add[11],z12.get()))
z12_send.place(x=370,y=270)

z13 = tk.Entry(window,width=10,font=('微软雅黑',16),show = None)
z13.place(x=20,y=270)
z13_send = tk.Button(window,text='13号登陆',fg='red',width=10,command=lambda :send(ip_add[12],z11.get()))
z13_send.place(x=120,y=270)

z14 = tk.Entry(window,width=10,font=('微软雅黑',16),show = None)
z14.place(x=270,y=310)
z14_send = tk.Button(window,text='14号登陆',width=10,command=lambda :send(ip_add[13],z14.get()))
z14_send.place(x=370,y=310)

z15 = tk.Entry(window,width=10,font=('微软雅黑',16),show = None)
z15.place(x=20,y=310)
z15_send = tk.Button(window,text='15号登陆',width=10,command=lambda :send(ip_add[14],z15.get()))
z15_send.place(x=120,y=310)

z16 = tk.Entry(window,width=10,font=('微软雅黑',16),show = None)
z16.place(x=270,y=350)
z16_send = tk.Button(window,text='16号登陆',width=10,command=lambda :send(ip_add[15],z16.get()))
z16_send.place(x=370,y=350)

z17 = tk.Entry(window,width=10,font=('微软雅黑',16),show = None)
z17.place(x=20,y=350)
z17_send = tk.Button(window,text='17号登陆',width=10,command=lambda :send(ip_add[16],z17.get()))
z17_send.place(x=120,y=350)

z18 = tk.Entry(window,width=10,font=('微软雅黑',16),show = None)
z18.place(x=270,y=390)
z18_send = tk.Button(window,text='18号登陆',width=10,command=lambda :send(ip_add[17],z18.get()))
z18_send.place(x=370,y=390)

z19 = tk.Entry(window,width=10,font=('微软雅黑',16),show = None)
z19.place(x=20,y=390)
z19_send = tk.Button(window,text='19号登陆',width=10,command=lambda :send(ip_add[18],z19.get()))
z19_send.place(x=120,y=390)

z20 = tk.Entry(window,width=10,font=('微软雅黑',16),show = None)
z20.place(x=270,y=430)
z20_send = tk.Button(window,text='20号登陆',width=10,command=lambda :send(ip_add[19],z20.get()))
z20_send.place(x=370,y=430)

z21 = tk.Entry(window,width=10,font=('微软雅黑',16),show = None)
z21.place(x=20,y=430)
z21_send = tk.Button(window,text='21号登陆',width=10,command=lambda :send(ip_add[20],z21.get()))
z21_send.place(x=120,y=430)

z22 = tk.Entry(window,width=10,font=('微软雅黑',16),show = None)
z22.place(x=270,y=470)
z22_send = tk.Button(window,text='22号登陆',width=10,command=lambda :send(ip_add[21],z22.get()))
z22_send.place(x=370,y=470)

z23 = tk.Entry(window,width=10,font=('微软雅黑',16),show = None)
z23.place(x=20,y=470)
z23_send = tk.Button(window,text='23号登陆',width=10,command=lambda :send(ip_add[22],z23.get()))
z23_send.place(x=120,y=470)

z24 = tk.Entry(window,width=10,font=('微软雅黑',16),show = None)
z24.place(x=180,y=510)
z24_send = tk.Button(window,text='24号登陆',width=10,command=lambda :send(ip_add[23],z24.get()))
z24_send.place(x=280,y=510)

# 客户端机器合集
all_client = {'1':z1,'2':z2,'3':z3,'4':z4,'5':z5,'6':z6,'7':z7,'8':z8,'9':z9,'10':z10,'11':z11,'12':z12,'13':z13,'14':z14,
              '15':z15,'16':z16,'17':z17,'18':z18,'19':z19,'20':z20,'21':z21,'22':z22,'23':z23,'24':z24}

all_sign = tk.Button(window,text='一键登陆',font=('微软雅黑',25),command=sign_all)
all_sign.place(x=150,y=570)

all_clean = tk.Button(window,text='全部清空',command=clean)
all_clean.place(x=350,y=615)

# 测试insert功能
# def insert_point():
#     var = z24.get()
#     z23.insert('end',var)

# test = tk.Button(window,text='test',command=insert_point)
# test.place(x=0,y=0)

messagelist = ['今天心情不好，自己动手去登陆吧', '让我自个崩溃，你去忙吧',
               '温馨提示：我崩溃了', '本程序随主人出走了，88']

tk.Label(window, text='作者：陈晓君', font=('微软雅黑',10)).place(x=410, y=680)
window.iconbitmap('2810_tk.ico')
namefile()
if (datetime.datetime.now() > datetime.datetime(2020,8,31,23,00)):
    tkinter.messagebox.showerror(title='程序崩溃,请联系作者',message=messagelist[datetime.datetime.now().minute % 4])
    os._exit(0)

window.mainloop()
