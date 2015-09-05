<%
Function RandomSay()
	n=Year(date())
	y=Month(date())
	r=Day(date())
	s=Hour(time())
	f=Minute(time())
	m=Second(time())
	if len(y)=1 then y="0" & y
	if len(r)=1 then r="0" & r
	if len(s)=1 then s="0" & s
	if len(f)=1 then f="0" & f
	if len(m)=1 then m="0" & m
	sj=s & ":" & f & ":" & m
	weekdate=weekday(date())	'星期值
	sjz=weekdate*10000+s*100+f	'夺宝大赛时间
	beginsj=weekdate*100000+s*1000+f*10+m '夺宝大赛准时开始时间620001  620005

	sjjh_name=Session("sjjh_name")
	sjjh_grade=Session("sjjh_grade")
	sjjh_jhdj=Session("sjjh_jhdj")
	nowinroom=session("nowinroom")

	'#####房间处理#####
	sjjh_roominfo=split(Application("sjjh_room"),";")
	chatinfo=split(sjjh_roominfo(nowinroom),"|")
	'开始随机事件
	if chatinfo(6)=0 then
		randomize()
		rnd1=int(rnd*300)+1
		'Response.Write "<script language=JavaScript>{alert("&rnd1&");}</script>"
		sayyq=""
		if beginsj>=621030 and beginsj<=621040  and application("sjjh_mpdz")<>1 then
		   Application.Lock
		   Application("sjjh_mpdz")=1              '1为开始，0为结束
		   Application.UnLock
		   Set conn=Server.CreateObject("ADODB.CONNECTION") 
		   Set rs=Server.CreateObject("ADODB.RecordSet") 
		   conn.open Application("sjjh_usermdb")
		   conn.Execute ("update 用户 set 门派杀人=0  ")
		   conn.close
		   set conn=nothing
		   sayyq="<bgsound src=wav/lunjian.mid loop=1><table width='85%' bgcolor='#CC6699' bordercolorlight='000000' bordercolordark='FFFFFF' border=1 cellspacing='0' cellpadding='2' align='center'><tr><td height='15'><div align='center'><font color='#FFFFFF' size='4'><font color=yellow face='Wingdings'>[</font><font color='yellow'><b>门 派 大 战 令</b></font><font color=yellow face='Wingdings'>]</font></font></div></td></tr><tr><td bgcolor='#6699CC'><div align='center'><font color='#000000' size='3'>"&Application("sjjh_user")&"宣布，门派大战现在开始...</font></div></td></tr></table>" 
		end if

		if beginsj>=720300 and beginsj<=720305 then
			yyyy="夺宝大赛现在已开始，夺宝大战的战况会不断在大厅进行报道..."
			sayyq="<bgsound src=wav/duobao.mid loop=1><table bgcolor=CC6699 width=85% align=center border=1 cellspacing=0 cellpadding='2' bordercolorlight=000000 bordercolordark=FFFFFF><tr><td align=center><font color=00FFFF><b>夺 宝 大 赛</b></font></td><tr><td bgcolor='#6699CC'><div align='center'><font size=-1>"&yyyy&"</font></div></td></tr></table>"
			rnd1=999999
		end if

		if rnd1>1 and (Minute(time())=30 and Second(time())<7 or Minute(time())=45 and Second(time())<7) then
			sayyq="<bgsound src=wav/02.wav loop=1><div align=center><img src='epk.gif' border='0'></div>"
		end if

		if rnd1<=7 and (Minute(time())>=20 and Minute(time())<30) then
			sayyq="<bgsound src=wav/duo.mid loop=1><table bgcolor=CC6699 width=85% align=center border=1 cellspacing=0 cellpadding='2' bordercolorlight=000000 bordercolordark=FFFFFF><tr><td align=center><font color=00FFFF><b>你 知 道 吗</b></font></td><tr><td bgcolor='#6699CC'><div align='center'><font color=#CC0000 size=-1>夺宝大赛每星期六晚20:30分准时在[高手房间]房间开始，战斗等级在30级以上、非官府、非出家人员均可参加，夺宝冠军<img src=sjfunc/guanjun.gif>将获得江湖至宝--紫金葫芦，修练完紫金葫芦后可以得到经验3000点，金币500个，体力、内力、武功上限各加1000点。</font></div></td></tr></table>"
		end if

		'以下是设置每小时前20分钟可以自动发钱，可以改的。
		if rnd1<=7 and Minute(time())<8 then
			useronlinename=Application("sjjh_useronlinename"&nowinroom)
			online=Split(Trim(useronlinename)," ",-1)
			x=UBound(online)
			Set conn=Server.CreateObject("ADODB.CONNECTION")
			conn.open Application("sjjh_usermdb")
			select case rnd1
				case 1,5
					randomize timer
					s=int(rnd*5000000)
					for i=0 to x
						conn.Execute "update 用户 set 银两=银两+" & s & " where 姓名='" & online(i) & "'"
					next
					sayyq="<bgsound src=wav/mail.wav loop=1><font color=green>【系统发钱】</font><B><font color=#ff0000>发钱喽！发钱喽！<img src=img/251.gif><img src=img/251.gif><img src=img/251.gif> 聊天室里的每个人发到"& s &"两。本江湖地址→www.happyjh.com别忘了叫你朋友一起来哦 ^_^</font></b><br><marquee width=100% behavior=alternate scrollamount=15></marquee><img src=img/022.gif><img src=img/022.gif><img src=img/022.gif><img src=img/022.gif></marquee>"
				case 2,7
					for i=0 to x
						conn.Execute "update 用户 set 金币=金币+5 where 姓名='" & online(i) & "'"
					next
					sayyq="<bgsound src=wav/mail.wav loop=1><font color=green>【系统金币】</font><B><font color=#ff0000>站长回首当年感谢大家对『快乐江湖』的支持和厚爱，心中高兴，特为聊天室的每个人发了金币5个！</font></b><br><marquee width=100% behavior=alternate scrollamount=15></marquee><img src=img/jinbi.gif><img src=img/jinbi.gif><img src=img/jinbi.gif><img src=img/jinbi.gif><img src=img/jinbi.gif></marquee>"
				case 3
					'randomize timer
					'mm=int(rnd*50)
					mm=50
					for i=0 to x
						conn.Execute "update 用户 set allvalue=allvalue+" & mm & " where 姓名='" & online(i) & "'"
					next
					sayyq="<bgsound src=wav/mail.wav loop=1><font color=green>【系统存点】</font><B><font color=#ff0000>站长路过此地，不忍看到各位大小虾如此苦练，特为聊天室的每个人发放经验" & mm & "点！</font></b>"
				case 4,6
					randomize timer
					s=int(rnd*100)
					for i=0 to x
						conn.Execute "update 用户 set 泡豆点数=泡豆点数+" & s & ",金=金+" & s & ",木=木+" & s & ",水=水+" & s & ",火=火+" & s & ",土=土+" & s & " where 姓名='" & online(i) & "'"
					next
					sayyq="<bgsound src=wav/mail.wav loop=1><font color=green>【系统发放】</font><B><font color=#ff0000>站长回首当年路经此地，给在座的各位发放豆子" & s & "点...呵呵，还有呢:为了鼓励大家修炼个人武功轩辕，给各位送上金、木、水、火、土属性各" & s & "点。</font></b>"
			end select
			conn.close
			set conn=nothing
		else
			select case rnd1
				case 8
				banner="感谢大家支持『快乐江湖』.回首当年将会在这套江湖上更用心的做到最完美.如果有任何意见可随时提出来.只要回首当年在都可以随时提出!本江湖的地址是www.happyjh.com别忘了叫你朋友一起来哦^_^;『快乐江湖』:做最美、最强的江湖！本江湖的地址是www.happyjh.com.如果您觉得这套程序不错.请告诉您的朋友.让他(她)们来此欢聚!这里是你我们的快乐家园!;『快乐江湖』所有用户注册时将拥有战斗等级10级.银两10000000.第一次登陆系统还自动发放新人费.让朋友们都有个较高的起点.使用更多的功能.但10级用户不能送钱不能转帐不能拍卖.所以请珍惜你的单个帐号.如果您的帐号超过3个.将随时删除.恕不相告.;『快乐江湖』由回首当年独立改版美化.使用权归回首当年个人所有.免费提供休闲娱乐、交友交流.注册用户初始化一级会员.三个月内若表现良好、等级达60级可继续升为二级会员.之后达90级可升为三级会员.达120级可升为四级会员.达300级可升为五级会员.如若不是则自动取消会员资格.到期前5天系统会提示您.谢绝不友之客！本江湖的地址是www.happyjh.com别忘了叫你朋友一起来哦^_^;站长友情提醒:如果您是学生.请注意不要影响休息.更不能影响学习.您可以挂网泡点.但不要在这花太多的时间和精力.『快乐江湖』的宗旨是休闲娱乐、交识网友.希望她能给您带来无限知识和欢乐!但不能因此而轻视了更重要的学习生活.切记!;在主页≮一线网络≯里设了『快乐江湖』的快速入口.另外还有以前的4个登陆页面.登陆时可以自选风格.也可以让系统随机.目前共有11种最新最酷的风格页面.进入后.主题音乐有35首.随机的.刷新就可以听到不同的乐曲了.有20几种鼠标样式以及各种漂亮的滚动条.你的浏览器IE版本是6.0的吗.如果不是.赶快升级吧;为了确保数据的安全.请大家及时申请密码保护.经常修改密码.不要随便把自己的帐号让他人来上.以免丢失.请文明聊天.本江湖的地址是www.happyjh.com别忘了叫你朋友一起来哦^_^;『快乐江湖』完全免费.在这里大家可以娱乐、可以交友、可以交流.所以请朋友们共同维护江湖秩序.遵守江湖规则.营造良好氛围.如果发现恶意扰乱秩序的.必杀！严重者封锁IP.请自觉合作!;『快乐江湖』每小时前10分钟系统发钱.最大值1000万两.泡点每分钟银两50000.武功内力50.存点10.站长也会时常给在线的朋友发放银两、金币、点数、体力、内力和武功等.希望您在场哦.^_^.请不要自己向站长要这要那的.或帮你改数据.因为那是没可能的事.切记！;江湖PK（比武）时间为每小时的后30分钟.在线人数必须超过3人才可以进行.一天的杀人数为10.夺宝一次算一次杀人.如果官府人员在线.比武时间到时请及时通知.大家可以尽情格斗.请新手和不愿意打架的开启练功保护.以防误杀.别伤和气哦^_^;凡战斗等级在30级以上、非官府、非出家人员均可参加夺宝大赛.夺宝大赛每星期六晚20:30整开始.要参加夺宝大赛的人员只能在大赛开始前10分钟内进入[高手房间]房间(即20:20-20:30之间进入).提前或超过此时间则无法进入夺宝之战房间.;夺宝大赛时禁止使用乾坤一掷、吸星大法、偷钱、拍卖、合体技、卡片、单挑、暂离、出家、还俗、同归于尽、点哑穴、配药、赠送.除此之外的功能均可以使用.动武时请点击聊天窗口右下角的夺宝(在非夺宝之战房间禁止使用).夺宝前10分钟进入[高手房间]房间.比赛开始后则无论你是否闭关保护.是否暂离都可以打人或被打.夺宝完成后.修练完紫金葫芦后可以得到的奖励有:经验3000点.金币500个.体力、内力、武功上限各加1000点.;江湖新手请注意→如果您是初次玩江湖或初次来本江湖.有什么疑问请不要急着问官府问别人.您可以先到论坛认真查看.在那里有本江湖的详细制度和规则.也有不少玩家的心得体验.相信会给您带来很多便利.增长您的江湖生命值!祝您好运!"
				banners=Split(Trim(banner),";",-1)
				total=UBound(banners)
				randomize timer
				x=int(rnd*(total+1))
				sayyq="<bgsound src=wav/gonggao.wav loop=1><table width=85% cellspacing=0 cellpadding=2 bgcolor=#009933 align=center bordercolorlight=000000 bordercolordark=FFFFFF border=1><tr><td align=center><font color=FFFFFF>★ 系 统 消 息 ★</font></td><tr><td bgcolor=CCCCFF> <div align='center'><font size=-1>"&banners(x)&"</font></div></td></tr></table>"
				case 9,7,5
				'取江湖笑话
				Set Conn=Server.CreateObject("ADODB.Connection")
				Set rs=Server.CreateObject("ADODB.RecordSet")
				Connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("ejoke.asp")
				Conn.Open connstr
				'取笑话的记录条数
				sql="select count(id) from jokelist"
				set rs=Conn.execute(sql)
				jokecount=rs(0)
				randomize
				rand=Int(jokecount * Rnd +1 )
				sql="select * from jokelist where id=" & rand & ""
				set rs=Conn.execute(sql)
				joke=rs("类型")&":"&rs("题目")&":"&rs("文章")
				rs.close
				set rs=nothing
				conn.close
				set conn=nothing
				joke=replace(joke,chr(13),"")
				joke=replace(joke,chr(10),"")
								Application.Lock
								Application("sjjh_joke")=joke
								Application.UnLock
								sayyq="<bgsound src=wav/haha.WAV loop=1><img src=joke.GIF><font color=009933>〖开心一笑〗</font><img src=jokepic.GIF><font color=666666>" & Application("sjjh_joke") & "</font>"			'聊天数据				
				case 11
						Application.Lock
						Application("sjjh_dalie")="老虎"
						Application.UnLock
						sayyq="<bgsound src=wav/hu.wav loop=1><font color=red>【消息】</font><b><font color=red>突然一只野兽<img src=img/laohu.gif>窜入江湖中伤人，请高手们快去打猎啊。</font></b>"
				case 12
						jstl=int(rnd*50000)+1000
						Application.Lock
						Application("sjjh_kl")=s
						Application.UnLock
						abc="<a href='dog.asp?tl="&Application("sjjh_kl")&"' target='d'><img src='img/dog.gif' border='0'></a>"
						sayyq="<bgsound src=wav/dog.wav loop=1><font color=red>【消息】</font><b><font color=red>突然间一声狗叫……“狗狗呀！”江湖里面怕狗的人尖叫，<bgsound src=wav/gougou.wav loop=1>大叫到不要咬我，不要咬我，我怕怕！一头狗狗闯进聊天室，狗狗体力：+"&jstl&"</font><b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
				case 13
						jstl=int(rnd*200000)+10000
						Application.Lock
						Application("sjjh_yb")=jstl
						Application.UnLock
						abc="<a href='yb.asp?tl="&Application("sjjh_yb")&"' target='d'><img src='img/251.GIF' border='0'></a>"
						sayyq="<bgsound src=wav/diaoxia.wav loop=1><font color=red>【消息】</font><b><font color=red>今天是什么日子，突然间从天上掉下来一个大元宝，价值：+"&jstl&"两!</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
				case 14
						jstl=int(rnd*50000)+10000
						Application.Lock
						Application("sjjh_kl")=jstl
						Application.UnLock
						abc="<a href='ws.asp?tl="&Application("sjjh_kl")&"' target='d'><img src='img/bt.GIF' border='0'></a>"
						sayyq="<bgsound src=wav/bsj.wav loop=1><font color=red>【消息】</font><b><font color=red>哇！『快乐江湖』来了一个武士，要找大家比武。武士体力："&jstl&"</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
				case 15
						jstl=int(rnd*500000)+10000
						Application.Lock
						Application("sjjh_yb")=jstl
						Application.UnLock
						abc="<a href='yb.asp?tl="&Application("sjjh_yb")&"' target='d'><img src='img/251.GIF' border='0'></a>"
						sayyq="<bgsound src=wav/diaoxia.wav loop=1><font color=red>【消息】</font><b><font color=red>今天真是好日子，"&Application("sjjh_user")&"中了E线风采头等奖，500万呀！给聊天室兄弟们一个大红包，价值：+"&jstl&"两!</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
				case 16
						jstl=int(rnd*50000)+1000
						Application.Lock
						Application("sjjh_kl")=jstl
						Application.UnLock
						abc="<a href='lj.asp?tl="&Application("sjjh_kl")&"' target='d'><img src='img/lj.GIF' border='0'></a>"
						sayyq="<bgsound src=wav/baguo.WAV loop=1><font color=red>【消息】</font><b><font color=red>八国联军入侵『快乐江湖』，江湖儿女们杀鬼子呀！冲呀~~~~~~八国联军人数："&jstl&"</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
				case 17
						jstl=int(rnd*50000)+1000
						Application.Lock
						Application("sjjh_kl")=jstl
						Application.UnLock
						abc="<a href='boy.asp?tl="&Application("sjjh_kl")&"' target='d'><img src='img/p42.gif' border='0'></a>"
						sayyq="<bgsound src=wav/FAQIAN.WAV loop=1><font color=red>【消息】</font><b><font color=red>一个衰哥听说这个江湖的美女特别多，跑到江湖想泡美女，大家准备好棒子打呀。打跑了官府有奖。。衰哥体力：+"&jstl&"</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
				case 18
						jstl=int(rnd*10000)+10000
						Application.Lock
						Application("sjjh_kl")=jstl
						Application.UnLock
						abc="<a href='mm.asp?tl="&Application("sjjh_kl")&"' target='d'><img src='img/mm.GIF' border='0'></a>"
						sayyq="<bgsound src=wav/wn2.wav loop=1><font color=red>【消息】</font><b><font color=red>咦哈~~~『快乐江湖』来了位大美女，看谁能抱得美女归，美女体力："&jstl&"</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
				case 19
						jstl=int(rnd*20000)+10000
						Application.Lock
						Application("sjjh_qi")=jstl
						Application.UnLock
						abc="<a href='qi.asp?tl="&Application("sjjh_qi")&"' target='d'><img src='img/tank.gif' border='0'></a>"
						sayyq="<bgsound src=wav/Bombs020.wav loop=1><font color=red>【消息】</font><b><font color=red>一把精制手枪闯进『快乐江湖』，中原武林人士看了目瞪口呆，也不知道这是什么东西。都先打了在说。。精制手枪体力：+"&jstl&"</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
				case 20
						jstl=int(rnd*10000)+10000
						Application.Lock
						Application("sjjh_kl")=jstl
						Application.UnLock
						abc="<a href='wn.asp?tl="&Application("sjjh_kl")&"' target='d'><img src='img/wn.GIF' border='0'></a>"
						sayyq="<bgsound src=wav/wn1.wav loop=1><font color=red>【消息】</font><b><font color=red>站长为女孩子们带来了『快乐江湖』最帅的领舞先生，举起你们的双手“摇起来~~~”，舞男体力："&jstl&"</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
				case 21
						jstl=int(rnd*10000)+10000
						Application.Lock
						Application("sjjh_ybb")=jstl
						Application.UnLock
						abc="<a href='ybb.asp?tl="&Application("sjjh_ybb")&"' target='d'><img src='img/251.GIF' border='0'></a>"
						sayyq="<bgsound src=wav/Phant008.wav loop=1><font color=red>【消息】</font><b><font color=red>E线特遣部队闯进『快乐江湖』，拿了扫水桶向所有聊天室的人挥洒，从聊天室每个人身上捞走：-"&jstl&"两!</font></b><br><marquee width=100% behavior=alternate scrollamount=15><img src='img/10.GIF' border='0'>"&abc&"</marquee>"
				case 22
					jstl=int(rnd*5)+1
					Application.Lock
					Application("sjjh_by")=jstl
					Application.UnLock
					abc="<a href='by.asp?tl="&Application("sjjh_by")&"' target='d'><img src='img/by.GIF' border='0'></a>"
					sayyq="<bgsound src=wav/diaoxia.wav loop=1><b><font color=red>【好消息】不知道谁的药掉在路上了：+"&jstl&"!谁抢到是谁的！</font></b><br><marquee width=100% behavior=alternate scrollamount=50>"&abc&"</marquee>"
				case 23
					jstl=int(rnd*100)+10
					Application.Lock
					Application("sjjh_yb")=jstl
					Application.UnLock
					abc="<a href='cd.asp?tl="&Application("sjjh_yb")&"' target='d'><img src='img/cd.GIF' border='0'></a>"
					sayyq="<bgsound src=wav/diaoxia.wav loop=1><b><font color=red>【抢存点】『快乐江湖』里飞来存点+"&jstl&"点!谁抢到是谁的！</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
				case 24
					jstl=int(rnd*50)+10
					Application.Lock
					Application("sjjh_yb")=jstl
					Application.UnLock
					abc="<a href='dd.asp?tl="&Application("sjjh_yb")&"' target='d'><img src='img/dd.GIF' border='0'></a>"
					sayyq="<bgsound src=wav/diaoxia.wav loop=1><b><font color=red>【抢豆点】『快乐江湖』里飞来豆点+"&jstl&"点!谁抢到是谁的！</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
				case 25
					jstl=int(rnd*2)+1
					Application.Lock
					Application("sjjh_yb")=jstl
					Application.UnLock
					abc="<a href='jk.asp?tl="&Application("sjjh_yb")&"' target='d'><img src='img/jk.GIF' border='0'></a>"
					sayyq="<bgsound src=wav/diaoxia.wav loop=1><b><font color=red>【抢金卡】『快乐江湖』里飞来金卡+"&jstl&"块!谁抢到是谁的！</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
				case 26
						jstl1=int(rnd*50000)+100
						Application.Lock
						Application("sjjh_yb1")=jstl1
						Application.UnLock
						abc="<a href='apple.asp?tl="&Application("sjjh_yb1")&"' target='d'><img src='img/apple.gif' border='0'></a>"
						sayyq="<bgsound src=wav/diaoxia.wav loop=1><font color=red>【消息】</font><b><font color=red>突然间从天上掉下来一个苹果，吃了加体力"&jstl1&"点!</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
				case 27
						rftl=int(rnd*50000)+10000
						Application.Lock
						Application("sjjh_rf")=rftl
						Application.UnLock
						abc="<a href='rf.asp?tl="&Application("sjjh_rf")&"' target='d'><img src='img/rf.GIF' border='0'></a>"
						sayyq="<bgsound src=wav/by.wav loop=1><b><font color=#FF0000>【捉人贩】</font></b><font color=#663300>人贩子:叶三娘“可恶的人贩子，把我那可怜的孩儿拐到哪里去啦？要是再找不到的话，我就见一个孩子杀一个。”人贩子体力：+"&rftl&"</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
				case 28
						piaoke=int(rnd*50000)+10000
						Application.Lock
						Application("sjjh_piaoke")=piaoke
						Application.UnLock
						abc="<a href='piaoke.asp?tl="&Application("sjjh_piaoke")&"' target='d'><img src='img/piaoke.GIF' border='0'></a>"
						sayyq="<bgsound src=wav/xiaotou.wav loop=1><b><font color=#FF0000>【抓小偷】</font></b><font color=#663300>任务开始:这年头江湖也不平静啊！居然有位贼头贼脑的小偷闯入江湖，谁把他逮住必有重谢！”</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
				case 29
						jstl=int(rnd*2)+1
						Application.Lock
						Application("sjjh_jinbi")=jstl
						Application.UnLock
						abc="<a href='jinbi.asp?tl="&Application("sjjh_jinbi")&"' target='d'><img src='img/jinbi.GIF' border='0'></a>"
						sayyq="<bgsound src=wav/diaoxia.wav loop=1><font color=red>【消息】</font><b><font color=red>今天是什么日子，突然间天上掉下+"&jstl&"个金币!</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
				case 30
					jstl=int(rnd*50000)+1000
					Application.Lock
					Application("sjjh_kl")=jstl
					Application.UnLock
					abc="<a href='DABIAN.ASP?tl="&Application("sjjh_kl")&"' target='d'><img src='img/ying1.gif' border='0'></a>"
					sayyq="<bgsound src=wav/niao.wav loop=1><font color=red>【消息】</font><b><font color=red>一只怪鸟从冥王星飞临江湖聊天室，太空鸟啊，一定会有好东西的，大家快打啊!怪鸟体力"&jstl&"</font></b><br><marquee width=100%  scrollamount=38>"&abc&"</marquee>"
				case 31
					jstl=int(rnd*0)+1
					Application.Lock
					Application("sjjh_kp")=jstl
					Application.UnLock
					abc="<a href='kp2.asp?tl="&Application("sjjh_kp")&"' target='d'><img src='img/kapian.GIF' border='0'></a>"
					sayyq="<bgsound src=wav/diaoxia.wav loop=1><font color=red>【消息】</font><b><font color=red>一张金光闪闪的卡片掉在江湖路上，各位大小虾快抢啊!</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
				case 32
					jstl=int(rnd*3)+1
					Application.Lock
					Application("sjjh_py")=jstl
					Application.UnLock
					abc="<a href='py2.asp?tl="&Application("sjjh_py")&"' target='d'><img src='img/py.GIF' border='0'></a>"
					sayyq="<bgsound src=wav/diaoxia.wav loop=1><font color=red>【消息】</font><b><font color=red>炼丹没炼成，配药散落一地，将配药丢在路上，谁捡到是谁的!</font></b><br><marquee width=100% behavior=alternate scrollamount=45>"&abc&"</marquee>"
				case 33
					jstl=int(rnd*0)+1
					Application.Lock
					Application("sjjh_lw")=jstl
					Application.UnLock
					abc="<a href='lw.asp?tl="&Application("sjjh_lw")&"' target='d'><img src='img/lw.jpg' border='0'></a>"
					sayyq="<bgsound src=wav/zhanfa.wav loop=1><font color=red>【消息】</font><b><font color=red>站长为五湖四海的朋友献礼了，祝愿大家天天开心、万事如意！</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
				case 34
					jstl=int(rnd*50000)+10000
					Application.Lock
					Application("sjjh_kl")=jstl
					Application.UnLock
					abc="<a href='liumang.asp?tl="&Application("sjjh_kl")&"' target='d'><img src='img/liumang.gif' border='0'></a>"
					sayyq="<bgsound src=wav/zhu.wav loop=1><font color=red>【消息】</font><b><font color=red>一个猪哥听说这个江湖的美女特别多，跑到江湖想泡美女，大家准备好棒子打呀。打跑了官府有奖。。流氓猪哥体力："&s+100&"</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
				case 35
						jstl1=int(rnd*50000)+100
						Application.Lock
						Application("sjjh_yb1")=jstl1
						Application.UnLock
						abc="<a href='bieshu.asp?tl="&Application("sjjh_yb1")&"' target='d'><img src='img/bieshu.gif' border='0'></a>"
						sayyq="<bgsound src=wav/diaoxia.wav loop=1><font color=red>【消息】</font><b><font color=red>『快乐江湖』为了感谢朋友们的支持和厚爱，送豪华别墅一栋，抢到加体力"&jstl1&"点!</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
				case 36
						jstl=int(rnd*50000)+10000
						Application.Lock
						Application("sjjh_kl")=jstl
						Application.UnLock
						abc="<a href='shu.asp?tl="&Application("sjjh_kl")&"' target='d'><img src='img/shu.gif' border='0'></a>"
						sayyq="<bgsound src=wav/shu.wav loop=1><font color=red>【消息】</font><b><font color=red>突然间江湖创入一只饿疯的老鼠……“饿鼠呀！”江湖里面怕鼠的人尖叫，大叫到不要咬我，不要咬我，我怕怕！！！一头饿鼠闯进聊天室，饿鼠体力：+"&jstl&"</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
				case 37
						jstl=int(rnd*2)+1
						Application.Lock
						Application("sjjh_jinbi")=jstl
						Application.UnLock
						abc="<a href='bxinshou.asp?tl="&Application("sjjh_jinbi")&"' target='d'><img src='img/xinshou.GIF' border='0'></a>"
						sayyq="<bgsound src=wav/xinshou.wav loop=1><font color=#000000>【照顾新手】</font><b><font color=red>有新朋友来『快乐江湖』了，站长说：介绍朋友来、照顾新人者奖:+"&jstl&"个金币!</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
				case 38,2
				'星期值*1000+小时值*100+分钟，星期五晚19点50则为6*10000+19*100+50=61950，星期五晚20:00则为6*10000+20*100+0=62000
					if weekdate=7 and sjz<72030 then
						Set conn=Server.CreateObject("ADODB.CONNECTION")
						Set rs=Server.CreateObject("ADODB.RecordSet")
						conn.open Application("sjjh_usermdb")
						rs.open "select * from 夺宝",conn
						if (rs("领取")=true or rs("宣布")=true) then
							conn.execute "update 夺宝 set 冠军=0,领取=false,修练次数=0,宣布=false,时间=now()"
						end if
						rs.close
						set rs=nothing
						conn.close
						set conn=nothing
						if sjz>=71820 and sjz<72023 then
							yyyy="江湖夺宝大赛即将开始，如有兴趣夺宝者请在<font color=red><b>20:20至20:30</b></font>之间进入夺宝房间[高手房间]，超过20:30后将不能再进入夺宝房间，夺宝大赛<b>20:30</b>分准时开始，请各位做好准备。比赛以在[高手房间]房间内只剩最后一个人时为结束标准，冠军为最后一个还在[高手房间]内的人。<font color=blue><b>奖品为：经验3000点，金币500个，并为冠军提高体力、内力、武功上限各1000点。</b></font>"
						else
							if sjz>=72020 and sjz<72030 then
								yyyy="江湖夺宝大赛现在可以<b><font color=blue>开始进入夺宝大战房间[高手房间]了</font></b>，请有意参赛者抓紧时间在<b><font color=blue>20:30之前</font></b>进入夺宝大战房间。夺宝大赛20:30分正式开始。在夺宝大赛时禁止使用卡片以及合体技，请各位在进入前做好充分准备。"
							else
								yyyy="今天是江湖夺宝大赛时间，有兴趣参加者可在每星期六晚上<b><font color=red>20:20分</font>至<font color=red>20：30分</font></b>之间进入[高手房间]，提前或超过这个时间将不能再进入[高手房间]。夺宝大赛以[高手房间]只剩最后一个人时为结束标志，最后这个人即为胜利者。在夺宝大赛时禁止使用卡片及合体技。奖品为经验3000点，金币500个，提高体力、内力、武功上限各1000点。"
							end if
						end if
						sayyq="<bgsound src=wav/tixing.wav loop=1><table bgcolor=CC6699 width=85% align=center border=1 cellspacing=0 cellpadding='2' bordercolorlight=000000 bordercolordark=FFFFFF><tr><td align=center><font color=00FFFF><b>夺 宝 大 赛</b></font></td><tr><td bgcolor='#6699CC'><div align='center'><font size=-1>"&yyyy&"</font></div></td></tr></table>"
					end if
				case 39
						jstl=int(rnd*5000)+1000
						Application.Lock
						Application("sjjh_kl")=jstl
						Application.UnLock
						abc="<a href='kl.asp?tl="&Application("sjjh_kl")&"' target='d'><img src='img/cins.GIF' border='0' width='60' height='80'></a>"
						sayyq="<bgsound src=wav/gui.wav loop=1><font color=red>【消息】</font><b><font color=red>突然间阴风陈陈……“僵尸呀！”一女子尖叫，一头僵尸闯进聊天室，僵尸体力：+"&jstl&"</font></b><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
				case 42,43,44
					'取江湖题库
					Set Conn=Server.CreateObject("ADODB.Connection")
					Set rs=Server.CreateObject("ADODB.RecordSet")
					Connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("questions.asp")
					Conn.Open connstr
					'取题库的记录条数
					sql="select count(id) from QuestLib "
					set rs=Conn.execute(sql)
					askcount=rs(0)
					randomize
					rand=Int(askcount * Rnd +1 )
					sql="select * from QuestLib where id=" & rand & ""
					set rs=Conn.execute(sql)
					ask=rs("类型")&":"&rs("问题")
					reply=rs("回答")
					rs.close
					set rs=nothing
					conn.close
					set conn=nothing
					ask=replace(ask,chr(13),"")
					ask=replace(ask,chr(10),"")
									Application.Lock
									Application("sjjh_ask")=ask
									Application("sjjh_reply")=reply
									Application("sjjh_askuser")="机器人"
									Application("sjjh_asksilver")=int(rnd*9999)+100000
									Application.UnLock
									sayyq="<bgsound src=wav/chuti.WAV loop=1><b>【系统提问】<font color=balck>" & Application("sjjh_ask") & "</font>正确答案是什么？[提问人]:<font color=red>【" & Application("sjjh_askuser")&"】</font></b><font color=green><b>奖励："&Application("sjjh_asksilver")&"两!</b></font>"			'聊天数据		
				case 10
					if Application("sjjh_baowu")>0 and sjjh_grade<6 then 
					   Set conn=Server.CreateObject("ADODB.CONNECTION") 
					   Set rs=Server.CreateObject("ADODB.RecordSet") 
					   conn.open Application("sjjh_usermdb") 
					   rs.open "select id,姓名,宝物,等级,门派,身份 from 用户 where 宝物='" & Application("sjjh_baowuname") & "'",conn 
					   if rs.eof or rs.bof  then 
						rs.close 
						rs.open "select 门派 from 用户 where 姓名='"&sjjh_name&"'" 
						if rs("门派")="出家" then 
						 sayyq1="江湖消息：##在江湖道上偶然见到了江湖至宝<img src=img/z1.gif><font color=red>"& Application("sjjh_baowuname")&"</font>，但由于自己是个出家人，出家人戒贪，只好看了一眼宝物，悻悻然而去~~~！" 
						 sayyq=Replace(sayyq1,"##",zj,1,3,1)  
						else 
						 conn.execute  "update 用户 set 保护=false,宝物修练=0,宝物='"& Application("sjjh_baowuname") &"' where 姓名='" & sjjh_name &"'" 
						 sayyq1="江湖消息：##现在拥有江湖至宝<img src=img/z1.gif><font color=red>"& Application("sjjh_baowuname")&"</font>" & Application("sjjh_baowusm") &"各位侠士互相争夺！" 
						 sayyq=Replace(sayyq1,"##",zj,1,3,1)  
						end if 
					   else 
						baouser=rs("姓名") 
						baowuid=rs("id") 
						if Instr(LCase(Application("sjjh_useronlinename"&nowinroom))," "&LCase(baouser)&" ")=0 then 
						 conn.execute  "update 用户 set 宝物='无',宝物修练=0 where 姓名='"& baouser &"'" 
						else 
						 sayyq1="江湖消息：江湖宝物<img src=img/z1.gif><font color=red>"& Application("sjjh_baowuname")&"</font>已被["&rs("门派")&"]"&rs("身份")&baouser&"夺走，各位大侠还不去抢..." 
						 sayyq=Replace(sayyq1,"##",zj)  
						end if 
					   end if 
					   rs.close 
					   set rs=nothing 
					   conn.close 
					   set conn=nothing 
					  end if
			end select			
		end if
	end if
	if sayyq<>"" Then
	
		act="消息"
		towhoway=0
		towho="大家"
		addwordcolor="660099"
		saycolor="008888"
		addsays="对"

		sayyq=replace(sayyq,"'","\'")
		sayyq=replace(sayyq,chr(34),"\"&chr(34))
		act="运气"
		saystr="<"&"script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & sjjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & sayyq & chr(39) &",0," & nowinroom & ");<"&"/script>"
		AddMsg SayStr
	end if
	
End Function
Function Yushu(a)
	Yushu=(a and 31)
End Function
Sub AddMsg(Str)
	Application.Lock()
	Application("SayCount")=Application("SayCount")+1
	i="SayStr"&YuShu(Application("SayCount"))
	Application(i)=Str
	Application.UnLock()
End Sub
%>