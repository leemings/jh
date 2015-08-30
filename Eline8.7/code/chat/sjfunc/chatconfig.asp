<%
sub dbdz(bb,cz)	'对是否在夺宝大战房间判断
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(session("nowinroom")),"|")
fjname=chatinfo(0)
if chatinfo(5)<>0 then
	erase sjjh_roominfo
	erase chatinfo
	Response.Write "<script language=JavaScript>{alert('提示：在『"&fjname&"』房间不可以使用"&cz&"！');}</script>"
	Response.End
end if
if bb=1 then
	if chatinfo(0)="高手房间" then
		erase sjjh_roominfo
		erase chatinfo
		Response.Write "<script language=JavaScript>{alert('提示：夺宝大赛时请重新点击窗口右下角功能区中的功能或点击夺宝，不可以使用"&cz&"！');}</script>"
		Response.End
	end if
end if
f=Minute(time())
if f<30 and fjname="快乐江湖" then
	erase sjjh_roominfo
	erase chatinfo
	Response.Write "<script language=JavaScript>{alert('提示："&fjname&"里只能在每小时的后30分钟才可以使用"&cz&"，想打架到『高手房间』房间去吧！"&fjname&"');}</script>"
	Response.End
end if
erase sjjh_roominfo
erase chatinfo
end sub

sub roompd(zsm)
	sjjh_roominfo=split(Application("sjjh_room"),";")
	chatinfo=split(sjjh_roominfo(session("nowinroom")),"|")
	fjname=chatinfo(0)
	erase chatinfo
	erase sjjh_roominfo
	if fjname<>"高手房间" then
		Response.Write "<script language=JavaScript>{alert('提示："& zsm &"使用只有在『高手房间』房间才可以！');}</script>"
		Response.End
	end if
	s=Hour(time())
	f=Minute(time())
	m=Second(time())
	weekdate=weekday(date())
	sjz=weekdate*10000+s*100+f
	'星期值*10000+小时值*100+分钟，星期五晚19点50则为6*10000+19*100+50=61950，星期五晚20:00则为6*10000+20*100+0=62000
	if sjz<72030 then
		Response.Write "<script language=JavaScript>{alert('提示：夺宝大赛还未开始，请稍后！');}</script>"
		Response.End
	end if
end sub
function zxrsbd(slname,slid)	'最后夺宝房间在线人数判断
	zxrs=application("sjjh_onlinelist"&nowinroom)
	 '开始统计在线人员数目
	rjsq=ubound(zxrs)
	if rjsq<=1 then
		erase zxrs
		Set conn1=Server.CreateObject("ADODB.CONNECTION")
		Set rs1=Server.CreateObject("ADODB.RecordSet")
		conn1.open Application("sjjh_usermdb")
		conn1.execute "update 夺宝 set 冠军="& slid &",时间=now(),领取=false,修练次数=0,宣布=true"
		set rs1=nothing
		conn1.close
		set conn1=nothing
		gg="<font color=blue>"& slname &"</font>经过艰苦奋战，终于夺得了<font color=red><b>夺宝大赛冠军</b></font>，站长"& application("sjjh_user") &"将武林至宝<font color=red><b>紫金葫芦</b></font>奖励给<font color=blue>"& slname &"</font><img src=sjfunc/guanjun.gif>，此宝需要修练可提升自已的各种上限及得到金币。"
		zxrsbd="<bgsound src=wav/gonggao.wav loop=1><table bgcolor=CC6699 width=85% align=center border=1 cellspacing=0 cellpadding='2' bordercolorlight=000000 bordercolordark=FFFFFF><tr><td align=center><font color=00FFFF><b>夺 宝 大 赛</b></font></td><tr><td bgcolor='#6699CC'><div align='center'><font size=-1>"&gg&"</font></div></td></tr></table>"
		exit function
	end if
	for i=1 to rjsq
		sfgf=split(zxrs(i),"|")
		if sfgf(0)<>slname and sfgf(2)<>"官府" then
			erase sfgf
			erase zxrs
			exit function
		end if
		erase sfgf
	next
	erase zxrs
	Set conn1=Server.CreateObject("ADODB.CONNECTION")
	Set rs1=Server.CreateObject("ADODB.RecordSet")
	conn1.open Application("sjjh_usermdb")
	conn1.execute "update 夺宝 set 冠军="& slid &",时间=now(),领取=false,修练次数=0,宣布=true"
	set rs1=nothing
	conn1.close
	set conn1=nothing
	gg="<font color=blue>"& slname &"</font>经过艰苦奋战，终于夺得了<font color=red><b>夺宝大赛冠军</b></font>，站长"& application("sjjh_user") &"将武林至宝<font color=red><b>紫金葫芦</b></font>奖励给<font color=blue>"& slname &"</font><img src=sjfunc/guanjun.gif>，此宝需要修练，可提升自已的各种上限及得到500个金币。"
	zxrsbd="<bgsound src=wav/gonggao.wav loop=1><table bgcolor=CC6699 width=85% align=center border=1 cellspacing=0 cellpadding='2' bordercolorlight=000000 bordercolordark=FFFFFF><tr><td align=center><font color=00FFFF><b>夺 宝 大 赛</b></font></td><tr><td bgcolor='#6699CC'><div align='center'><font size=-1>"&gg&"</font></div></td></tr></table>"
end function
function exitltzxpd(inroom) '退出聊天室判断
	sjjh_roominfo=split(Application("sjjh_room"),";")
	chatinfo=split(sjjh_roominfo(inroom),"|")
	fjname=chatinfo(0)
	erase sjjh_roominfo
	erase chatinfo
	if fjname<>"高手房间" then
		exitltzxpd=""
		exit function
	end if
	session("nowinroom")=0
	Set conn1=Server.CreateObject("ADODB.CONNECTION")
	Set rs1=Server.CreateObject("ADODB.RecordSet")
	conn1.open Application("sjjh_usermdb")
	rs1.open "select * from 夺宝",conn1,2,2
	xb=rs1("宣布")
	rs1.close
	if xb then
		set rs1=nothing
		conn1.close
		set conn1=nothing
		exitltzxpd=""
		exit function
	end if
	zxrs=application("sjjh_onlinelist"&inroom)
	 '开始统计在线人员数目
	rjsq=ubound(zxrs)
	jsq=0
	sname=""
	if rjsq<=1 then
		set rs1=nothing
		conn1.close
		set conn1=nothing
		exitltzxpd=""
		exit function
	end if
	for i=1 to rjsq
		sfgf=split(zxrs(i),"|")
		if sfgf(0)<>session("sjjh_name") and sfgf(2)<>"官府" then
			sname=sfgf(0)
			jsq=jsq+1
		end if
		if jsq>1 then
			erase zxrs
			set rs1=nothing
			conn1.close
			set conn1=nothing
			exitltzxpd=""
			exit function
		end if
		erase sfgf
	next
	erase zxrs
	if sname<>"" then
		rs1.open "select id from 用户 where 姓名='"&sname&"'",conn1
		if not (rs1.eof or rs1.bof) then
			slid=rs1("id")
			rs1.close
			conn1.execute "update 夺宝 set 冠军="& slid &",时间=now(),领取=false,修练次数=0,宣布=true"
		else
			rs1.close
			set rs1=nothing
			conn1.close
			set conn1=nothing
			exitltzxpd=""
			exit function
		end if
	else
		set rs1=nothing
		conn1.close
		set conn1=nothing
		exitltzxpd=""
		exit function
	end if
	set rs1=nothing
	conn1.close
	set conn1=nothing
	gg="<font color=blue>"& sname &"</font>经过艰苦奋战，终于夺得了<font color=red><b>夺宝大赛冠军</b></font>，站长"& application("sjjh_user") &"将武林至宝<font color=red><b>紫金葫芦</b></font>奖励给<font color=blue>"& sname &"</font><img src=sjfunc/guanjun.gif>，此宝需要修练，可提升自已的各种上限及得到500个金币。"
	exitltzxpd="<bgsound src=wav/gonggao.wav loop=1><BR><BR><table bgcolor=CC6699 width=85% align=center border=1 cellspacing=0 cellpadding='2' bordercolorlight=000000 bordercolordark=FFFFFF><tr><td align=center><font color=00FFFF><b>夺 宝 大 赛</b></font></td><tr><td bgcolor='#6699CC'><div align='center'><font size=-1>"&gg&"</font></div></td></tr></table>"
end function
sub dbxx(xx)
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
	t="<font class=t>(" & sj & ")</font>"
	says="<font color=CC0000><b>【夺宝消息】</b></font>" & xx			'聊天数据
	says=replace(says,chr(39),"\'")
	says=replace(says,chr(34),"\"&chr(34))
	act="消息"
	towhoway=0
	towho="大家"
	addwordcolor="660099"
	saycolor="008888"
	addsays="对"
	saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & sjjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ",0);<"&"/script>"
	addmsg saystr
end sub
Function Yushu(a)
	Yushu=(a and 31)
End Function
Sub AddMsg(Str)
	Application.Lock()
	Application("SayCount")=Application("SayCount")+1
	i="SayStr"&YuShu(Application("SayCount"))
	Application(i)=Str
	Application.UnLock()
end sub
%>