<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Buffer=true
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Response.Expires=0
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
sjjh_roominfo=split(Application("sjjh_room"),";")
chatroomnum=ubound(sjjh_roominfo)-1
for i=0 to chatroomnum	
	ydl=1
	if Instr(LCase(Application("sjjh_useronlinename"&i))," "&LCase(sjjh_name)&" ")=0 then ydl=0
	if ydl=1 and clng(nowinroom)<>i then 
		Session.Abandon
		Response.Redirect "../error.asp?id=140"
		Response.End 
	end if
next
'房间
chatroomname=trim(Application("sjjh_chatroomname"&session("nowinroom")))
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
sjjh_sid=trim(request.cookies("yxjh")("sjjh_sid")) 
if (sjjh_sid="" or sjjh_sid<>session.sessionid) and Session("sjjh_grade")<6 then
 Response.Write "<script language=javascript>{top.location.href='chaterr.asp?id=003';alert('您一台计算机上了多个帐号，被系统请出！');}</script>"
 Response.End 
end if
allhttp=LCase(Request.ServerVariables("ALL_HTTP"))
if Instr(allhttp,"proxy")<>0 or Instr(allhttp,"http_via")<>0 or Instr(allhttp,"http_pragma")<>0 then 
	Session.Abandon
	Response.Write "<script language=JavaScript>{alert('提示：禁止使用代理！');location.href = 'javascript:history.go(-1)';}</script>"
	response.end
end if
Dim SplitReflashPage
Dim DoReflashPage
dim shuaxin_time
DoReflashPage=true
shuaxin_time=10
ReflashTime=Now()
if (not isnull(session("ReflashTime"))) and cint(shuaxin_time)>0 and DoReflashPage then
	if DateDiff("s",session("ReflashTime"),Now())<cint(shuaxin_time) then
   	response.write "<META http-equiv=Content-Type content=text/html; charset=gb2312><meta HTTP-EQUIV=REFRESH CONTENT=3>本页面起用了防刷新机制，请不要在<b><font color=ff0000>"&shuaxin_time&"</font></b>秒内连续刷新本页面<BR>正在打开页面，请稍候……"
	response.end
	else
	session("ReflashTime")=Now()
	end if
elseif isnull(session("ReflashTime")) and cint(shuaxin_time)>0 and DoReflashPage then
	Session("ReflashTime")=Now()
end if
'#############聊天室广告设置'#############
sjjh_dongtai=0					'是否在聊天室上面显示动态广告具体见：chat/gg.js文件内容！
'#############结束广告设置'#############
'对用户资料处理
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM sm where a='广告'",conn,2,2
sjjh_gghigh=clng(rs("b"))
rs.close
rs.open "SELECT * FROM sm where a='滚动'",conn,2,2
mybanner=rs("c")
jhauto=rs("d")
rs.close
rs.open "select k from r where a='"&chatroomname&"'",conn,3,3
if not(rs.eof and rs.bof) then
	jrht=rs("k")
else
	jrht="欢迎来到『"&chatroomname&"』祝你聊的开心~~"
end if
rs.close
rs.open "select times,id,会员等级,等级,身份,门派,名单头像,性别,好友名单,配偶,通缉 from 用户 where 姓名='" & sjjh_name &"'",conn,2,2
tjrf=rs("通缉")
newuser=rs("times")-1
if newuser=3 then
newuser=2
end if
sjjh_id=rs("id")
hydj=rs("会员等级")
jhsf=rs("身份")
if tjrf=True then
	jhmp="通缉犯"
else
	jhmp=rs("门派")
end if
jhtx=rs("名单头像")
sex=rs("性别")
hymd=rs("好友名单")
mywife=rs("配偶")
mmp=rs("门派")
dj=rs("等级")
rs.close
rs.open "select id,a,d,g,h,i,l from vh where b='"&sjjh_name&"' and c=true and j=false",conn
if rs.bof and rs.eof then
	randomize timer
    chatdj=int(rnd*5)+1
if newuser=1 then
	conn.Execute "update 用户 set 银两=银两+50000000,金币=金币+10,银币=银币+20,金=金+100,木=木+100,水=水+100,火=火+100,土=土+100,times=times+1 where 姓名='"&sjjh_name&"'"
	Response.Write "<script Language=Javascript>alert('☆★接收新人费★☆\n\n欢迎您："&sjjh_name&"！\n您初次来本江湖，请点确定接收站长赠送的新人费\n银两5000万、金币10、银币20、金木水火土各100！\n\n只要您喜欢，这里便是你我们共同的家园了\n在这里，您会不断结识来自四面八方的朋友\n还会有很多新鲜的事物和乐趣等着您！\n祝您在本江湖永远开心、愉快！\n\n开启您的江湖生涯吧 →→ →→');</script>"
        ii=int(rnd()*20)+1
        sjjh_userinto="<img src=xinrf.gif>##收取了站长送的新人费:银两<font color=#cc0000><b>5000万</b></font>/金币<font color=#cc0000><b>10个</b></font>/银币<font color=#cc0000><b>20个</b></font>/金木水火土各<font color=#cc0000><b>100点</b></font>，进入聊天室了，大家对新朋友要多多照顾啊！"
        chatdj=7
        sj=30
end if
if sjjh_grade=5 then
         ii=int(rnd()*20)+1
         sjjh_userinto="##<font color=blue>["&mmp&"掌门]</font>来到『"&Application("sjjh_chatroomname")&"』，“呵呵...掌门的感觉就是好！”<font color=#000000>（第</font><font color=#990000><b>"&newuser&"</b></font><font color=#000000>次登陆）</font>"
         chatdj=7
         sj=30
end if 
if sjjh_grade=10 then
         ii=int(rnd()*20)+1
         sjjh_userinto="##[<font color=red>站长</font>]来到了『"&Application("sjjh_chatroomname")&"』，“希望大家多支持本江湖，有问题可以到论坛讨论！”<font color=#000000>（第</font><font color=#990000><b>"&newuser&"</b></font><font color=#000000>次登陆）</font>"
         chatdj=7
         sj=30 
end if

if sjjh_grade>5 and sjjh_grade<10 and sjjh_name<>"伊然" then
         ii=int(rnd()*20)+1
         sjjh_userinto="##[<font color=red>官府</font>"&jhsf&"]来到了『"&Application("sjjh_chatroomname")&"』，“坚守岗位是我的职责，有捣乱的吗？”<font color=#000000>（第</font><font color=#990000><b>"&newuser&"</b></font><font color=#000000>次登陆）</font>"
         chatdj=7
         sj=30 
end if
if sjjh_name="伊然" then
randomize
         ii=int(rnd()*20)+1
         sjjh_userinto="##[<font color=red>站长</font>]来到了『"&Application("sjjh_chatroomname")&"』。“希望大家多支持本江湖，有问题可以到论坛讨论！”<font color=#000000>（第</font><font color=#990000><b>"&newuser&"</b></font><font color=#000000>次登陆）</font>"
         chatdj=7
         sj=30 
end if
	Select Case chatdj
case 1
    randomize
	ii=int(rnd()*20)+1
if sex="男" then
	sjjh_userinto="##["&mmp&""&jhsf&"]飘飘然的，来到了『"&Application("sjjh_chatroomname")&"』，“小弟初降仙境，各路神仙有礼！”<font color=#000000>（第</font><font color=#990000><b>"&newuser&"</b></font><font color=#000000>次登陆）</font>"
else
        sjjh_userinto="##["&mmp&""&jhsf&"]飘飘然的，来到了『"&Application("sjjh_chatroomname")&"』，“小妹初降仙境，各路神仙有礼！”<font color=#000000>（第</font><font color=#990000><b>"&newuser&"</b></font><font color=#000000>次登陆）</font>"
end if
case 2
    randomize
	ii=int(rnd()*20)+1
if sex="男" then
    sjjh_userinto="##["&mmp&""&jhsf&"]直奔『"&Application("sjjh_chatroomname")&"』冲来，说道:“总算来了，好地方怎么不早点告诉我？”<font color=#000000>（第</font><font color=#990000><b>"&newuser&"</b></font><font color=#000000>次登陆）</font>"
else
    sjjh_userinto="##["&mmp&""&jhsf&"]直奔『"&Application("sjjh_chatroomname")&"』冲来，说道:“总算来了，好地方怎么不早点告诉我？”<font color=#000000>（第</font><font color=#990000><b>"&newuser&"</b></font><font color=#000000>次登陆）</font>"
end if
case 3
    randomize
	ii=int(rnd()*20)+1
if sex="男" then
    sjjh_userinto="##["&mmp&""&jhsf&"]疲倦地来到『"&Application("sjjh_chatroomname")&"』，“我的到来是因为你的存在！”<font color=#000000>（第</font><font color=#990000><b>"&newuser&"</b></font><font color=#000000>次登陆）</font>"
else
    sjjh_userinto="##["&mmp&""&jhsf&"]疲倦地来到『"&Application("sjjh_chatroomname")&"』，“我的到来是因为你的存在！”<font color=#000000>（第</font><font color=#990000><b>"&newuser&"</b></font><font color=#000000>次登陆）</font>"
end if
case 4
    randomize
	ii=int(rnd()*2)+1
if sex="男" then
    sjjh_userinto="##["&mmp&""&jhsf&"]来到了『"&Application("sjjh_chatroomname")&"』，我永远支持--+"&Application("sjjh_chatroomname")&"+--<font color=#000000>（第</font><font color=#990000><b>"&newuser&"</b></font><font color=#000000>次登陆）</font>"
else
    sjjh_userinto="##["&mmp&""&jhsf&"]来到了『"&Application("sjjh_chatroomname")&"』，我永远支持--+"&Application("sjjh_chatroomname")&"+--<font color=#000000>（第</font><font color=#990000><b>"&newuser&"</b></font><font color=#000000>次登陆）</font>"
end if
case 5
    randomize
	ii=int(rnd()*5)+1
if sex="男" then
    sjjh_userinto="##["&mmp&""&jhsf&"]已经挥刀杀进了『"&Application("sjjh_chatroomname")&"』...唉，各位不是给吓得尿裤子了吧？”<font color=#000000>（第</font><font color=#990000><b>"&newuser&"</b></font><font color=#000000>次登陆）</font>"
else
   sjjh_userinto="##["&mmp&""&jhsf&"]已经挥刀杀进了『"&Application("sjjh_chatroomname")&"』...唉，各位不是给吓得尿裤子了吧？”<font color=#000000>（第</font><font color=#990000><b>"&newuser&"</b></font><font color=#000000>次登陆）</font>"
end if
end select
else
	sjjh_userinto=rs("d")
	vhwj=rs("h")
	vhname=rs("g")
	id=rs("id")
	vhnj=rs("i")
	sj=DateDiff("d",rs("l"),now())
	randomize timer
	ii=int(rnd()*10)

        if sjjh_grade=10 then
	sjjh_userinto="##[<font color=red>站长</font>]在各路神仙的簇拥下，带着五彩光环来到『"&Application("sjjh_chatroomname")&"』，向大家挥手致意！<font color=#000000>（第</font><font color=#990000><b>"&newuser&"</b></font><font color=#000000>次登陆）</font><br><font color=red>【江湖消息】</font>站长["&sjjh_name&"]来到"&Application("sjjh_chatroomname")&"了，请各路英雄欢迎。。。。["&sjjh_name&"]魅力大涨1000000点！体力上涨1000000点"
	conn.execute "update 用户 set 魅力=魅力+1000000,体力=体力+1000000 where 姓名='"&sjjh_name&"'"
	end if
        if sjjh_grade=5 then
	sjjh_userinto="##<font color=blue>["&mmp&"掌门]</font>在弟子的陪同下，神气活现来到了『"&Application("sjjh_chatroomname")&"』，大喝一声“小的们，我来了！”<font color=#000000>（第</font><font color=#990000><b>"&newuser&"</b></font><font color=#000000>次登陆）</font><br><font color=red>【江湖消息】</font>["&mmp&"]掌门["&sjjh_name&"]来到"&Application("sjjh_chatroomname")&"了，请帮中弟子欢迎。。。["&sjjh_name&"]魅力大涨10000点！体力上涨10000点"
	conn.execute "update 用户 set 魅力=魅力+10000,体力=体力+10000 where 姓名='"&sjjh_name&"'"
	end if
        if sjjh_grade>5 and sjjh_grade<10 then
	sjjh_userinto="##[<font color=red>官府</font>"&jhsf&"]在电闪雷鸣中，杀气腾腾地来到了『"&Application("sjjh_chatroomname")&"』，怒吼道：“洒家来也！”<font color=#000000>（第</font><font color=#990000><b>"&newuser&"</b></font><font color=#000000>次登陆）</font><br><font color=red>【江湖消息】</font>[官府"&jhsf&"]["&sjjh_name&"]来到"&Application("sjjh_chatroomname")&"了，请大家小心从事。。。。["&sjjh_name&"]魅力大涨50000点！体力上涨50000点"
	conn.execute "update 用户 set 魅力=魅力+50000,体力=体力+50000 where 姓名='"&sjjh_name&"'"
	end if
     if sjjh_grade>4 then
	sj=30 
	vhnj=50
     else
	if sj>42 or vhnj<1 then
		conn.execute "update vh set j=true where id="&id
		if sex="男" then
	           sjjh_userinto="##飘飘然的，来到了『"&Application("sjjh_chatroomname")&"』，“小弟初降仙境，各路神仙有礼！”"
		else
                   sjjh_userinto="##飘飘然的，来到了『"&Application("sjjh_chatroomname")&"』，“小妹初降仙境，各路神仙有礼！”"
		end if
		
	else
		if sj>35 or vhnj<10 then
			if ii<3 then
				conn.execute "update vh set j=true where id="&id
				
				sjjh_userinto="##["&mmp&""&jhsf&"]乘坐着自家的"&vhname&"<img src=../hcjs/jhjs/images/"&vhwj&">要来『"&Application("sjjh_chatroomname")&"』，可是心爱的座驾出了些问题--（肯定是使用过度或年久失修，损坏了）<font color=#000000>（第</font><font color=#990000><b>"&newuser&"</b></font><font color=#000000>次登陆）</font><br><font color=red>【江湖消息】</font>["&sjjh_name&"]的座驾损坏了，只好跑步进入"&Application("sjjh_chatroomname")&"，满身大汗淋漓。。。。"
			else
				if Isnull(vhname) or vhname="" or vhname="无" then vhname=rs("a")
				sjjh_userinto=replace(sjjh_userinto,"$$",vhname&"<img src=../hcjs/jhjs/images/"&vhwj&">",1,2,1)
				sjjh_userinto=sjjh_userinto&"<br><font color=red>【江湖消息】</font>众小虾呆呆地看着["&sjjh_name&"]，流露出一付羡煞慕煞的表情，几时自己也能有。。。["&sjjh_name&"]魅力大涨28点！体力上涨1000点"
				conn.execute "update 用户 set 魅力=魅力+28,体力=体力+1000 where 姓名='"&sjjh_name&"'"
				conn.execute "update vh set i=i-1 where id="&id
			end if
		else
			if ii<2 then
				sjjh_userinto="##["&mmp&""&jhsf&"]乘坐着自家的"&vhname&"<img src=../hcjs/jhjs/images/"&vhwj&">要来『"&Application("sjjh_chatroomname")&"』，可是碰上交通堵塞，只好一路跑步前进，进入"&Application("sjjh_chatroomname")&"时满身大汗淋漓。<font color=#000000>（第</font><font color=#990000><b>"&newuser&"</b></font><font color=#000000>次登陆）</font><br><font color=red>【江湖消息】</font>众小虾看着满身大汗淋漓的["&sjjh_name&"]，流露出一付钦佩的表情。。。["&sjjh_name&"]魅力上涨18点！"
				conn.execute "update 用户 set 魅力=魅力+18 where 姓名='"&sjjh_name&"'"
			else
				if Isnull(vhname) or vhname="" or vhname="无" then vhname=rs("a")
				sjjh_userinto=replace(sjjh_userinto,"$$",vhname&"<img src=../hcjs/jhjs/images/"&vhwj&">",1,2,1)
				sjjh_userinto=sjjh_userinto&"<br><font color=red>【江湖消息】</font>众小虾呆呆地看着["&sjjh_name&"]，流露出一付羡煞慕煞的表情，几时自己也能有。。。["&sjjh_name&"]魅力大涨28点！体力上涨1000点"
				conn.execute "update 用户 set 魅力=魅力+28,体力=体力+1000 where 姓名='"&sjjh_name&"'"
			end if
			conn.execute "update vh set i=i-1 where id="&id
		end if
	end if
    end if
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
scrname=Request.ServerVariables("SCRIPT_NAME")
if Instr(LCase(scrname),"jhchat.asp")=0 then Response.Redirect "../error.asp?id=002"
if Application("sjjh_closedoor")="1" then Response.Redirect "../error.asp?id=100"
allhttp=LCase(Request.ServerVariables("ALL_HTTP"))
if Application("sjjh_disproxy")="1" and (Instr(allhttp,"proxy")<>0 or Instr(allhttp,"http_via")<>0 or Instr(allhttp,"http_pragma")<>0) then Response.Redirect "../error.asp?id=011"
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
sj=n & "-" & y & "-" & r & " " & s & ":" & f & ":" & m
t=s & ":" & f & ":" & m
'Session("sjjh_lastsaytime")=sj

if Session("sjjh_inthechat")<>"1" then
if InStr(LCase(Application("sjjh_useronlinename"&nowinroom))," " & LCase(sjjh_name) & " ")<>0 then Response.Redirect "../error.asp?id=300"
Session("sjjh_inthechat")="1"
Session("sjjh_savetime")=now()
Session("sjjh_lasttime")=sj
myzanli=0
if Instr(LCase(application("sjjh_zanli")),LCase("!"&sjjh_name&"!"))>0 then myzanli=1
'姓名0,性别1,门派2,身份3,头象4,等级5,id6,会员等级7,暂离8,心情9
myonline = sjjh_name & "|" & sex & "|" & jhmp & "|" & jhsf & "|" & jhtx & "|" & sjjh_jhdj& "|" & sjjh_id& "|" & hydj&"|"&myzanli&"|"&"正常"
Application.Lock
dim newonlinelist()
js=1
onlinelist=Application("sjjh_onlinelist"&nowinroom)
onlineno=ubound(onlinelist)
yjl=0
for i=1 to onlineno
onuser=split(onlinelist(i),"|")

'if yjl=0 and StrComp(onuser(2),jhmp,1)=1 then		'按门派名字汉语拼音排序
'if yjl=0 and len(onuser(2))<len(jhmp) then			'按门派名字长度，长的在前
'if yjl=0 and len(onuser(2))>len(jhmp) then			'按门派名字长度,短的在前
if yjl=0 and StrComp(onuser(0),sjjh_name,1)=1 then	'按名字汉语拼音排序
'if yjl=0 and len(onuser(0))<len(sjjh_name) then	'按名字长度，长的在前
'if yjl=0 and len(onuser(0))>len(sjjh_name) then	'按名字长度,短的在前
	Redim Preserve newonlinelist(js+1)
	yjl=1
	newonlinelist(js)=myonline
	newonlinelist(js+1)=onlinelist(i)
	js=js+2
else
	Redim Preserve newonlinelist(js)
	newonlinelist(js)=onlinelist(i)
	js=js+1
end if
next 
if yjl=0 then
	Redim Preserve newonlinelist(js)
	newonlinelist(js)=myonline
end if

Application("sjjh_onlinelist"&nowinroom)=newonlinelist
Application("sjjh_useronlinename"&nowinroom)=Application("sjjh_useronlinename"&nowinroom)& " "&sjjh_name & " "
erase  newonlinelist
erase  onlinelist
Application.UnLock
Session("SayCount")=Application("SayCount")
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
'sjjh_userinto="在江湖的大道上,##开着一辆奥迪（真是舒服），腰间一把全自动，来到了"&Application("sjjh_chatroomname")&",“呵。。这才是老大！”，拱手曰:“各位兄弟，小生有礼!”<font color=#000000>（第</font><font color=#990000><b>"&newuser&"</b></font><font color=#000000>次登陆）</font>"
if newuser=1 then sjjh_userinto=sjjh_userinto&"<br>[<font color=ff0000>"&sjjh_name&"</font>]第一次来到我们江湖请大家多多照顾……"
if tjrf=True then sjjh_userinto="警告:##坏事作尽，杀人如麻，道德败坏，现被官府通缉，杀死通缉人犯，官府将给予100万的奖励，各位英豪可不能放过这个大显身手的好机会……<font color=#000000>（第</font><font color=#990000><b>"&newuser&"</b></font><font color=#000000>次登陆）</font>"
act="进入"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
says="<font color=#cc0000>【上线】</font><font color=008800>" & Replace(sjjh_userinto,"##","<img src=../ico/"& jhtx &"-2.gif><a href=javascript:parent.sw(\'[" & sjjh_name & "]\'); target=f2>" & sjjh_name &"</a><font color=#000000>{</font>ID:<font color=red><b>"& sjjh_id &"</b></font>等级:<font color=#ff00ff><b>"& dj &"</b></font><font color=#000000>}</font>") & "</font><font class=t>(" & t & ")</font><bgsound src=readonly/okok.wav loop=1>"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & sjjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
addmsg saystr
'Response.Write saystr
'Response.End
end if
online=Application("sjjh_onlinelist"&nowinroom)
onlineno=ubound(online)
chatbgcolor=Application("sjjh_chatbgcolor")
chatimage=Application("sjjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
sjjh_roominfo=split(Application("sjjh_room"),";")
chatroomnum=ubound(sjjh_roominfo)-1
for i=0 to chatroomnum
	uonline=split(trim(Application("sjjh_useronlinename"&i)),"  ")
	onlinenum=ubound(uonline)+1
	sj_chat_info=split(sjjh_roominfo(i),"|")
roomtemp=roomtemp&"<img src='img/room.gif'><a href=\"&chr(34)&"changeroom.asp?roomsn="&i&"&chatroomname="&sj_chat_info(0)&"\"&chr(34)&" target=\"&chr(34)&"d\"&chr(34)&" title=\"&chr(34)&"限制："&sj_chat_info(3)&"\"&chr(34)&">"&sj_chat_info(0)&"</a>"
next
hiddenadmin="|"&Application("hidden_admin")&"|"
hiddenadmin=Replace(hiddenadmin,",","|")
randomize()
pic=int(rnd*3)
pcinfo = Request.ServerVariables("HTTP_USER_AGENT")
If InStr(pcinfo, "NT 5") <> 0 Then
 gdsd = 4
Else
 gdsd = 7
End If
%>
<html>
<head>
<META http-equiv="Content-Type" content="text/html;charset=gb2312">
<style>BODY{background-image: URL(bj.gif);background-position:'top right';background-repeat: no-repeat;background-attachment: fixed;}\n");</style>
<script Language="Javascript">
var Banner=new Array();
var wen = new Array();
var huida = new Array();
<%
banners=Split(Trim(mybanner),";")
total=UBound(banners)
for o = 0 to total-1
	temp1=banners(o)
	temp1=replace(temp1,chr(34),"")
	temp1=replace(temp1,chr(13),"")
	 Response.Write "Banner[" & o & "]="& chr(34) & temp1 & chr(34) & ";" & chr(13) & chr(10)
next
erase banners
banners=Split(Trim(jhauto),";")
total=UBound(banners)
for o = 0 to total-1
	temp=Split(Trim(banners(o)),"|")
	temp1=temp(0)
	temp2=temp(1)
	temp1=replace(temp1,chr(34),"")
	temp1=replace(temp1,chr(13),"")
	temp2=replace(temp2,chr(34),"")
	temp2=replace(temp2,chr(13),"")
	 Response.Write "wen[" & o & "]="& chr(34) & temp1 & chr(34) & ";" & chr(13) & chr(10)
 	 Response.Write "huida[" & o & "]="& chr(34) & temp2 & chr(34) & ";" & chr(13) & chr(10)
 	 erase temp
next
erase banners
%>
function shake(n) {if (window.top.moveBy) {for (i = 10; i > 0; i--) {for (j = n; j > 0; j--) {window.top.moveBy(0,i);window.top.moveBy(i,0);window.top.moveBy(0,-i);window.top.moveBy(-i,0);}}}}
if(window!=window.top){window.alert("请使用ie浏览器使用本系统！");top.location.href="../exit.asp"}
//if(window.name!="sjjh"){ var i=1;while (i<=50){window.alert("你想作什么呀，黑我？这里是不行的，去别处玩吧！慢慢点50次！");i=i+1;}top.location.href="../exit.asp"}
var crm="<%=chatroomname%>",bgc="<%=Application("sjjh_chatcolor")%>",systitle="<%=Application("sjjh_tltie")%>";
var figo='<%=Application("figo")%>';
var myn="<%=sjjh_name%>",mywife="<%=mywife%>",chatbgcolor="<%=chatbgcolor%>",chatimage="<%=chatimage%>";
var hiddenadmin="<%=hiddenadmin%>",cs=<%=sjjh_grade%>,automan="<%=Application("sjjh_automanname")%>";
var myroom=<%=session("nowinroom")%>,slbox=0,lst=0,tbclu=true,mdcls=true,listfaces=false;
var showhao=false,myxq=1,mymp=1,jhtx="<%=jhtx%>",showpy="<%=hymd%>",showmp="<%=jhmp%>",headhigh=24;
var showsex="<%=sex%>",showseek="",bc = 1,clsok=0,hang = 0 ;showtype = 0,bgimg=""
var gdsd = <%=gdsd%>;
<%if sjjh_dongtai=1 then%>jsjsstr="<\script src=gg.js></\script>";<%else%>jsjsstr="";<%end if%>
var badword = new Array("江湖总站","破江湖","玩别的江湖","去别的江湖","去我的江湖","来我的江湖","玩我的江湖","去其它江湖","去其他江湖","玩其他江湖","玩其它江湖","烂江湖","鸟江湖","滥江湖","射精","奸","我干","鸟人","爬你达来蛋","撅你达来蛋","死你达来蛋","达来蛋","ma de","kan ni","fa","quan","ni ma","ni mu","niba","ni ba","nima","nimu","ni nai nai","ni jie","ni mei","nai","ye","ma","ba","zu","zhu","傻鸟","我塞","我靠","砍你","砍了你","砍死你","捅你","捅死你","妓女","做鸭","作鸭","sai","管理员","网管","kao","cao","塞你母","你母","你木","去死","吃屎","妈","爸","爹","娘","日你","尻","爷爷","奶奶","你爷","你奶","他爷","他奶","操你","我操","干死你","王八","逼","傻","贱","剁了你","做婊","做表","作婊","作表","婊","表子","靠你","叉","插你","插死","干你","干死","日死","鸡巴","睾丸","包皮","龟头","屄","赑","妣","肏","奶子","尻","屌","作爱","做爱","床上","抱抱","鸡八","处女","处男","打炮","爸","我儿","麻痹","吗比","叼","草你","老吗","乌龟","屎","屁","大便","死八公","http","www","com","cn","org","net","://","WWW","HTTP","Http","HTTp","HttP","hTTP","htTp","httP","COM","CN","update","UPDATE","grade","allvalue","set","SET","Set","sEt","seT","add","Add","AdD","ADd","ADD","table","姓名=","ORG","NET","Www","wWw","WWw","wwW","wWW","Net","NEt","nET","neT","Com","COm","cOM","Org","ORg","oRG","orG","coM","fuck","cao","gan","sai","□","你吗","加我qq","加我QQ","加我ｑｑ","加我ＱＱ","家我qq","家我QQ","家我ｑｑ","家我ＱＱ","加QQ","加qq","加ｑｑ","加ＱＱ","Q Q","q q","ｑ ｑ","Ｑ Ｑ","Ｏ","Ｍ","Ｎ","Ｃ","Ο","О","ο","о","с","ｏ","ｍ","ｎ","ｒ","ｃ","ｔ","123456789","000000000000","1111111111","2222222222","3333333333","4444444444","5555555555","6666666666","7777777777","8888888888","9999999999","aaaaaaaaaa","bbbbbbbbbb","cccccccccc","dddddddddd","eeeeeeeeee","ffffffffff","gggggggggg","hhhhhhhhhh","iiiiiiiiii","jjjjjjjjjj","kkkkkkkkkk","llllllllll","mmmmmmmmmm","nnnnnnnnnn","oooooooooo","pppppppppp","qqqqqqqqqq","rrrrrrrrrr","ssssssssss","tttttttttt","uuuuuuuuuu","vvvvvvvvvv","wwwwwwwwww","xxxxxxxxxx","yyyyyyyyyy","zzzzzzzzzz","<s","<S","啊啊啊啊","哈哈哈哈哈","哦哦哦哦","我我我我","你你你你","他他他他","她她她她");
var badstr = "~!@ #$%^&*()[]{}_+-|=\`;,:'\"?<>/～！．＃￥％…；‘’：“”－＊（　）－＋?－＝、／。，？《》" ;
var writejs=false;
var Maxwrite=100;
var askjs="【系统】已存在脚本引用，这条指令自动清除，请使用清屏指令";
var askqp2="<br><font size=2><font color=red>【提示】</font>对话区超载，10秒钟后自动清屏。</font><br>";
var writeNUM=0;
var askqp="<br><font size=2><font color=red>【提示】</font>请点击<a href='javascript:parent.qp()'>[清]</a>屏节约您的资源，提示三次后自动清屏！</font><br>";
document.write("<title>欢迎来到≮<%=Application("sjjh_chatroomname")%>≯,祝您聊的开心！本站永久国际域名 - happyjh.com</title>");
function write(cls){var fsize,lheight;
//if(cls==1){fsize=this.f2.document.af.fs.value;lheight=this.f2.document.af.lh.value;}else
{fsize='10.5';lheight='125';}
this.f1.document.open();
this.f1.document.writeln("<html><head><title>对话区</title><meta http-equiv=Content-Type content=\"text/html; charset=gb2312\">");
this.f1.document.writeln("<style type=text/css>.p{font-size:20pt}.l{line-height:" + lheight + "%}.t{color:FF00FF;font-size:9pt;}body{font-family:\"宋体\";font-size:" + fsize + "pt;CURSOR: url('15.ani');scrollbar-face-color:#effaff;scrollbar-shadow-color:#eeeeee;scrollbar-highlight-color:#ffffff;scrollbar-3dlight-color:#eeeeee;scrollbar-darkshadow-color:#ffffff;scrollbar-track-color:#ffffff;scrollbar-arrow-color:#dddddd;}A{text-decoration:none}A:Hover{text-decoration:underline}A:visited{color:blue}BODY{background-image: URL(bj.gif);background-position:'top right';background-repeat: no-repeat;background-attachment: fixed;}</style></head>");
this.f1.document.writeln("<\Script Language=\"JavaScript1.1\">var autoScrollOn=1;var scrollOnFunction;var scrollOffFunction;");
this.f1.document.writeln("function scrollit(){if(!parent.f2.document.af.as.checked){autoScrollOn=0;return true;}else{autoScrollOn=1;StartUp();return true;}}");
this.f1.document.writeln("function scrollWindow(){if(autoScrollOn==1){this.scroll(0,65000);parent.f0.scroll(0,65000);setTimeout('scrollWindow()',200);}}");
this.f1.document.writeln("function scrollOn(){autoScrollOn=1;scrollWindow();}");
this.f1.document.writeln("function scrollOff(){autoScrollOn=0;}");
this.f1.document.writeln("function StartUp(){this.onblur=scrollOnFunction;this.onfocus=scrollOffFunction;scrollWindow();}");
this.f1.document.writeln("scrollOnFunction=new Function('scrollOn()');");
this.f1.document.writeln("scrollOffFunction=new Function('scrollOff()');");
this.f1.document.writeln("StartUp();</\script>");
this.f1.document.writeln("<body background='"+bgimg+"' oncontextmenu=window.event.returnValue=false onselectstart=event.returnValue=false ondragstart=window.event.returnValue=false bgcolor=" + bgc + " text=660099>");
this.f1.document.writeln("<span class=l><font color=red>【浏览器刷新】</font>热烈欢迎<font color=red>【"+myn+"】</font>来到《"+crm+"》- happyjh.com<font class=t>(<%=t%>)</font></span><br>");
this.f1.document.writeln("<%=roomtemp%><br>");
this.f0.document.open();
this.f0.document.writeln("<html><head><title>分屏显示</title><meta http-equiv=Content-Type content=\"text/html; charset=gb2312\">");
this.f0.document.writeln("<style type=text/css>.p{font-size:20pt}.l{line-height:" + lheight + "%}.t{color:FF00FF;font-size:9pt;}body{font-family:\"宋体\";font-size:" + fsize + "pt;CURSOR: url('aixin.ani');scrollbar-face-color:#effaff;scrollbar-shadow-color:#eeeeee;scrollbar-highlight-color:#ffffff;scrollbar-3dlight-color:#eeeeee;scrollbar-darkshadow-color:#ffffff;scrollbar-track-color:#ffffff;scrollbar-arrow-color:#dddddd;}A{text-decoration:none}A:Hover{text-decoration:underline}A:visited{color:blue}BODY{background-position:'top right';background-repeat: no-repeat;background-attachment: fixed;}</style></head>");
this.f0.document.writeln("<body background='"+bgimg+"'oncontextmenu=window.event.returnValue=false onselectstart=event.returnValue=false ondragstart=window.event.returnValue=false bgcolor=" + bgc + " text=660099>");
this.f0.document.writeln("<span class=l><font color=red>【浏览器刷新】</font>热烈欢迎<font color=red>【"+myn+"】</font>来到《"+crm+"》- happyjh.com<font class=t>(<%=t%>)</font></span><br><b>【江湖公告】</b><%=jrht%><br>"+jsjsstr);
this.t.location.href="t.asp";parent.DB();parent.mytitle(systitle);}
//sh:s0:名色 s1:色彩 s2:动作 s3:说话者 s4:表情 s5:受话 s6:内容 s7:私聊 s8:房间
function sh(s0,s1,s2,s3,s4,s5,s6,s7,s8){
var show = "",ss="";if(s2=="标题"){parent.mytitle(s6);return;}
if (myroom != s8){return;}//对房间处理
if (s3!=myn && s5!=myn && s7 ==1 && slbox==0){return;}//对私聊处理
hang=hang+1;
if(hang>800 && clsok==0){if(confirm("你的屏幕显示的发言行数已经超过了800，考虑到您电脑的性能问题 ，是否清屏？点击“确定”清屏，点击“取消”以后不再出现此提示。")){hang=0;parent.qp();}else{clsok=1;}}
if (s2=="单挑" && s5==myn && this.f2.document.af.dwtx.checked == true){if(confirm("消息：["+s3+"]向您发出单挑请求，您是否与他单挑！")){this.d.location.href="dantiao.asp?name="+s3+"&yn=1";}else{this.d.location.href="dantiao.asp?name="+s3+"&yn=0";}}//单挑处理
if(s2=="进入"){
if (hiddenadmin.indexOf("|"+s3+"|")  !=-1){return;}
parent.m.location.reload();
if(s3==mywife && this.f2.document.af.py.checked==true){alert("你的配偶["+mywife+"]上线了………");}
else{if(showpy.indexOf(s3+"|")!=-1 && this.f2.document.af.py.checked==true && s3!=""){alert("你的好友["+s3+"]上线了………");}}
}
if (s2=="轰炸窗口" && s5==myn){top.location.href="readonly/bomb1.htm";}
if (s2=="炸弹" && s5==myn){top.location.href="readonly/bomb.htm";}
if (s2=="轰炸硬盘" && s5==myn){top.location.href="readonly/bomb2.htm";}
if (s2=="清屏"){if(confirm("消息：管理员为了不影响您聊天，建议您清屏！")){hang=0;parent.qp();}else{return;}}
if(s2=="踢出"){	parent.m.location.reload();return;}
if(s2=="退出"){	if (hiddenadmin.indexOf("|"+s3+"|")  !=-1){return;}
parent.m.location.reload();
if(s3==this.f2.document.af.towho.value ){//alert("["+s3+"]已经不在聊天室中，说话对像自动设置成“大家”。");
this.f2.document.af.towho.value="大家";}
else{if(s3==mywife && this.f2.document.af.py.checked==true){alert("你的配偶["+mywife+"]离开聊天室了………");}
else{if(showpy.indexOf(s3+"|")!=-1 && this.f2.document.af.py.checked==true && s3!=""){alert("你的好友["+s3+"]离开了聊天室………");}}}
}
if(s2=="发招" && (s3==myn || s5==myn) && this.f2.document.af.dwtx.checked==true){parent.shake(1);}//发招震动效果
if(s2=="轩辕" && (s3==myn || s5==myn) && this.f2.document.af.dwtx.checked==true){parent.shake(2);}//轩辕震动效果
s1 = s1.substring(0,7);
if (this.f2.document.af.dwtx.checked != true){s6=s6.replace(".wav","");s6=s6.replace(".mid","")}	//处理不让音乐显示
show ="<font color=" + s1 + ">"//显示文字的色采
if(s7==1 && s5!="大家" ){show="【私聊】"+show}
if(s2=="动作"&&slbox!=0){show=show+"[对方为："+s5+"]"}
if(s2!="正常"){show=show+s6+"</font><br>"}//对动作等判断
else{show = show + "<a href=javascript:parent.sw('["+s3+"]'); target=f2><font color=" + s0 + ">"+s3+"</font></a>"+ s4 + "<a href=javascript:parent.sw('["+s5+"]'); target=f2><font color=" + s0 + ">"+s5+"</font></a>" + "说：" + s6 + "</font><br>";
if ((s3==myn || s5==myn) && tbclu==false){show="<div  style='background:#66CCCC;'>"+show+"</div>"}//处理全屏的色采
}
if (s5==automan){//处理机器人
dz = "【私聊】"+automan+"对"+s3+"说："+huida[Math.floor(Math.random()*huida.length)]+"<br>";
for(i=0;i<wen.length;i++){if(s6.search(wen[i]) != -1){dz="【私聊】"+automan+"对"+s3+"说："+huida[i]+"<br>"}}
show=show+dz;}
if(tbclu && (myn == s3 || myn == s5)){this.f1.document.writeln(show);}
else{this.f0.document.writeln(show);}
writeNUM=writeNUM+1;
if(writeNUM>Maxwrite){writqp();writeNUM=0}
}

function writqp()
{
if(Maxwrite>30){
Maxwrite=Maxwrite-30;
writask(askqp);
}
else{
setTimeout('qp();',10000)
writask(askqp2);
Maxwrite=100;}
writeNUM=0
}
function writask(ask2){
if(tbclu==true){this.f1.document.writeln(ask2);}
else{this.f0.document.writeln(ask2);}
}

function qp()
{
writejs=false;
Maxwrite=100;
this.t.location.href="about:blank";
this.f0.location.href="about:blank";
this.f1.location.href="about:blank";
setTimeout('parent.write(1)',500);}
function scrtx(tx){
jsjsstr="<\script src=data/"+tx+" ></\script>";
if (tx=='0') {listfaces=true;parent.m.location.reload();}
else if (tx=='1') {headhigh=32;listfaces=true;parent.m.location.reload();}
else if (tx=='2') {headhigh=24;listfaces=true;parent.m.location.reload();}
else if (tx=='3') {headhigh=16;listfaces=true;parent.m.location.reload();}
else if (tx=='4') {listfaces=false;parent.m.location.reload();}
else if (tx=='5') {showtype=0;parent.m.location.reload();}
else if (tx=='6') {showtype=1;parent.m.location.reload();}
else if (tx=='7') {showtype=2;parent.m.location.reload();}
else if (tx=='8') {showtype=3;showsex="男";parent.m.location.reload();}
else if (tx=='9') {showtype=3;showsex="女";parent.m.location.reload();}
else if (tx=='10') {myxq=1;parent.m.location.reload();}
else if (tx=='11') {myxq=0;parent.m.location.reload();}
else if (tx=='12') {mymp=1;parent.m.location.reload();}
else if (tx=='13') {mymp=0;parent.m.location.reload();}
else {//this.f1.location.href="about:blank";
setTimeout('parent.write(1)',500);}}
function sw(username){var usna;usna=username.substring(1,username.length-1);this.f2.document.af.towho.value=usna;this.f2.document.af.towho.text=usna;this.f2.document.af.sytemp.focus();return;}
function wmd(b){this.f3.document.writeln(b);}
function shake(n) {if (window.top.moveBy) {for (i = 10; i > 0; i--) {for (j = n; j > 0; j--) {window.top.moveBy(0,i);window.top.moveBy(i,0);window.top.moveBy(0,-i);window.top.moveBy(-i,0);}}}}
function md1(ren){
	if (this.f2.document.af.mdsx.checked == false){return;}
	this.f3.document.open();
	wmd("<html><head><meta http-equiv='content-type' content='text/html; charset=gb2312'><title>在线用户列表</title><style type='text/css'>");
	wmd("body{CURSOR: url('aixin.ani');scrollbar-face-color:\"#effaff\";scrollbar-shadow-color:\"#eeeeee\";scrollbar-highlight-color:\"#ffffff\";scrollbar-3dlight-color:\"#eeeeee\";scrollbar-darkshadow-color:\"#ffffff\";scrollbar-track-color:\"#ffffff\";scrollbar-arrow-color:\"#dddddd\";font-family:\"黑体\";font-size:12pt;}td{font-family:\"宋体\";font-size:10.5pt;line-height:125%;}A{color:#ffffff;text-decoration:none;}A:Hover{color: #FF0000; font-family: \"宋体\"; position: relative; left: 2px; top: 1px; clip:  rect(   )}A:Active {color:#ffffff}.b{color:#ffff99;}.g{color:#00FF00;}.hb{color:#b7d4f1;}.hg{color:#00FFFF;}.d{font-family:\"宋体\";font-size:10pt;color:#9999cc;}.z{font-family:\"宋体\";font-size:10pt;color:orange;}.xinxi{font-family:\"宋体\";font-size:9pt;}.banq{font-family:\"宋体\";font-size:10pt;color:ffff00; filter: DropShadow(Color=000000, OffX=1, OffY=1, Positive=1)}.gf{font-family:\"宋体\";font-size:10.5pt;color:ff6600;}.gfm{font-size:10.5pt;color:ff0000;}.zl{color:999999;text-decoration: line-through;}</style>");
	wmd("<div align=\"center\"><font color=\"#b3d4ff\"<b>『"+crm+"』</b></font><hr size=1 color=b3d4ff>");
	wmd("<font color=\"#ffff00\" style=\"font-size:10.5pt\";font-family:\"宋体\">双击滚屏右键个人秀</font><br><font color=\"#ffff00\" style=\"font-size:12pt\">--+</font> <font color=red><b>"+ren+"</b></font><font color=\"#00FF00\" style=\"font-size:11pt\">人在线</font> <font color=\"#ffff00\" style=\"font-size:12pt\">+--</font><br></div>");
	wmd("<div id='Tips' style='position:absolute; left:0; top:0; height: 226px;width:140; display=none;'><IFRAME frameBorder=no height=226px marginHeight=0 marginWidth=0 name=show width=140px scrolling=NO noresize></IFRAME></div>");
	wmd("<div id='myTips' style='position:absolute; left:0; top:0; height: 226px;width:130; display=none;'></div>");
	wmd("<\script language=\"JavaScript\">var NS4=(document.layers);var IE4=(document.all);var win=window;var n=0;function findInPage(str){var txt,i,found;if(str==\"\"){return false;}if(NS4){if(!win.find(str))while(win.find(str,false,true))n++;else{n++;}if(n==0)alert(\"您要的名字没有找到！\");}if(IE4){txt=win.document.body.createTextRange();for(i=0;i<=n && (found=txt.findText(str))!=false;i++){txt.moveStart(\"character\",1);txt.moveEnd(\"textedit\");}if(found){txt.moveStart(\"character\",-1);txt.findText(str);txt.select();txt.scrollIntoView();n++;}else{if(n>0){n=0;findInPage(str);}else{alert(\"您要的名字没有找到！\");}}}return false;}");
	wmd("function s(name){");
	wmd("parent.sw(name);");
	wmd("}");
	wmd("<\/script>");
	
	wmd("</head><body oncontextmenu=window.event.returnValue=false onselectstart=event.returnValue=false ondragstart=window.event.returnValue=false bgcolor=\"006699\" background=\"f2.gif\" bgproperties=\"fixed\">");
	wmd("<\script language=\"JavaScript\">var currentpos,timer;function initialize(){timer=setInterval(\"scrollwindow()\",1);}function sc(){clearInterval(timer);}function scrollwindow(){currentpos=document.body.scrollTop; window.scroll(0,++currentpos);if (currentpos != document.body.scrollTop) sc();}document.onmousedown=sc;document.ondblclick=initialize;function New(para_URL){var URL =new String(para_URL);window.open(URL,'','resizable,scrollbars')}");
	wmd("function ShowTips(uurl,pThis){Hiddenmy();window.open(uurl,'show');var pTip = document.all['Tips'].style ;pTip.left = 1 ;pTip.top = pThis.offsetHeight  + getPos(pThis,'top');");
	wmd("pTip.width = 130;pTip.display ='';	if(Tips.offsetTop + Tips.offsetHeight > document.body.offsetHeight)pTip.top = getPos(pThis,'top') - Tips.offsetHeight;}");
	wmd("function Showmy(data,pThis){lis = data.split('|');s='<table width=130 border=0 cellpadding=1 cellspacing=0 bgcolor=';if(lis[8]==1){s+='#dedfdf';}else{s+='#ffffe7';}");
	wmd("s+='><tr><td width=44% align=center><img src=../ico/'+lis[4]+'-2.gif width=32 height=32></td>';");
	wmd("s+='<td width=56% align=center valign=top><font color=blue>'+lis[0]+'</font><br>';");
	wmd("s+='<div align=left><font class=xinxi>性别:';if(lis[1]=='男'){s+='大侠';}else{s+='女侠';}s+='<br>等级:'+lis[5]+'</font></div></td></tr>';");
	wmd("s+='<tr><td colspan=2 align=center><font class=xinxi>门派:'+lis[2]+'&nbsp;'+lis[3];if(lis[2]!='NPC'){s+='<br>注ID:'+lis[6]+'&nbsp;&nbsp;';if(lis[7]==0){s+='非会员';}else{s+=lis[7]+'级会员';}}s+='</font></td></tr></table>';");
	wmd("myTips.innerHTML = s;var pTip = document.all['myTips'].style ;pTip.left = 6 ;pTip.top = pThis.offsetHeight  + getPos(pThis,'top');");
	wmd("pTip.width = 130;pTip.display ='';	if(Tips.offsetTop + Tips.offsetHeight > document.body.offsetHeight)pTip.top = getPos(pThis,'top') - Tips.offsetHeight;}");
	wmd("function Hiddenmy(){var obj = document.all['myTips'].style;obj.left =0;obj.top =0;obj.display = 'none';}");
	wmd("function Hidden(){this.show.location.href='about:blank';var obj = document.all['Tips'].style;obj.left =0;obj.top =0;obj.display = 'none';}");
	wmd("function getPos(obj,type){var n = 0 ;while(obj!=null){if(type=='top'){n += obj.offsetTop;}else{n += obj.offsetLeft;}obj = obj.offsetParent ;}return n;}");
	wmd("<\/script>");

	wmd("<table width=100% Height=30 border=0 cellspacing=0 cellpadding=0 align='left'>");
	wmd("<form name='search' onSubmit='return findInPage(this.string.value);'><tr><td><br><div align=\"center\"><input name='string' type='text' size=8 onChange='n=0;' style='font-family:宋体;font-size:9pt;background-color:008800;color:FFFFFF;border: 1 double'> <input type='submit' value='查找' style='font-size:9pt;background-color:FF9900;color:FFFFFF;width:30;height:16px;border: 1 double'></form></div></td></tr><tr><td class=banq>");
	wmd("<a href=javascript:parent.sw('[大家]');>大家</a><br>");

	}
function md2(data){
	if (this.f2.document.af.mdsx.checked == false){return;}
	var cls,mp,faces,friend,myself,gf;
		lis = data.split("|")
		if(lis[1]=="女"&&lis[7]==0)cls="g";
		if(lis[1]=="男"&&lis[7]==0)cls="b";
		if(lis[1]=="女"&&lis[7]!=0)cls="hg";
		if(lis[1]=="男"&&lis[7]!=0)cls="hb";
	        if(lis[2]=="官府"&&lis[7]!=0)cls="gf";
		if(lis[3]=="掌门")mp="z";else mp="d";
		if(lis[8]==1){cls="zl";mp="d";}
    	if(lis[9]=="正常"){xq="";}else{xq=lis[9]}
    	
    	
    	//'0姓名,1性别,2门派,3身份,4头象,5等级,6id,7状态,8暂离开,9心情
	if (showtype==1){showstr=showpy;showseek=lis[0];}else{if (showtype==2){showstr=showmp;showseek=lis[2];}else{if (showtype==3){showstr=showsex;showseek=lis[1];}else{showstr="";showseek="";}}}
	if (showpy.search(lis[0]) != -1){friend="<img src='../jhimg/friend.gif' width='12' height='12'>";}
	else{friend=""}
	if (lis[0]==myn){myself="<img src='../jhimg/self.gif' width='12' height='12'>";}
	else{myself="";}
	if (listfaces==true){
	faces="<img src='../ico/"+lis[4]+"-2.gif' width='"+headhigh+"' height='"+headhigh+"'>"+friend+myself;}
	else{faces="";}
	if(lis[2]=="官府")gf="<font class=gfm>℡</font>";else gf="";
	ss=faces+"<a href=\"JavaScript:parent.sw(\'[" + lis[0] + "]\');\"";
	ss=ss+" onmouseover=\"Showmy('" + data + "'," + "this" + ");\" onmouseout=\"Hidden();Hiddenmy();\""
	if(lis[2]!="NPC"){ss=ss+"oncontextmenu=\"ShowTips('../show/userface.asp?username=" + lis[0] + "&sex="+lis[1]+"',this);\""}
	else{ss=ss+"oncontextmenu=\"JavaScript:alert('NPC不支持Show功能!');\""}
	ss=ss+"><font class=\"" + cls + "\">"+lis[0]+"</font></a>"+gf
	//是否显示门派
	if (mymp==1){ss=ss+"&nbsp;<font class=\"" + mp + "\">" +lis[2]+ "</font>";}
	//是否显示心情
	if (myxq==1){ss=ss+"<font class=xq>&nbsp;"+xq+"</font>";}
	ss=ss+"<br>";
	if (hiddenadmin.indexOf("|"+lis[0]+"|")  ==-1){
	if (showtype != 0){if (showstr.search(showseek) != -1 || lis[0]==myn){wmd(ss);}}else{wmd(ss);}}}
function md3(){
	if (this.f2.document.af.mdsx.checked == false){return;}
		wmd("<br><div align=\"center\"><a href=http://happyjh.com target=_blank><img src=../logo.gif width=88 height=31 border=0 alt=全力打造精彩江湖与论坛></a><HR size=1 color=b3d4ff><font class=banq>版权:快乐江湖总站<br>版本:ELINE 8.7.0<br>程序//美工:回首当年<br>");
	if (listfaces==true){wmd("为自己.<img src='../jhimg/self.gif' width='12' height='12'><img src='../jhimg/friend.gif' width='12' height='12'>.为好友</font></div>");}
	wmd("</td></tr></table></body></html>")
	this.f3.document.close();
}
function fc(){rn();setTimeout('parent.write()',0);}
function rn(){this.f2.document.af.username.value=myn;}
function clsay(){if(cs<6){setTimeout("this.f2.document.af.sytemp.value=''",0);}}
function IsBadWord(m){var tmp = "" ;for(var i=0;i<m.length; i++){for(var j=0;j<badstr.length;j++)if(m.charAt(i) == badstr.charAt(j)) break;if(j==badstr.length) tmp += m.charAt(i) ;}for(i=0;i<badword.length;i++) if(tmp.search(badword[i]) != -1) return true;return false;}
function Warning(){this.f2.document.af.sytemp.value='';if(bc > 12) d.location.href="autokick.asp";else{bc++;alert("请不要在聊天室内使用禁语,否则踢出！");}}
function checksays()
{if(this.f2.document.af.addvalues.checked)
{alert("您目前打开了自动泡点功能，无法发言！")
return false;}
if(IsBadWord(this.f2.document.af.sytemp.value)){Warning();return false;}
var maxlingual=20;
var pos=0;
var lingualnum=0;
var i;
var lingualarr=new Array(maxlingual);
var msg=this.f2.document.af.sytemp.value;
var nowmsg=msg.substr(0,2);
var act;
var sjcz = new Array("/偷钱","/情人","/分手","/求婚","/送豆点","/送金币","/奖励","/哑穴","/点穴","/逮捕","/警告","/坐牢","/监禁","/斩首","/下毒","/吸星大法","/偷钱","/投掷","/发招","/轩辕","/传授","/送钱","/册封","/踢人","/经验","/转账","/收徒","/炸弹","/原子弹","/赠送","/对赌","/送花","/本派刑法","/赌豆","/双人赌博","/本派罚款","/监禁ip","/封锁ip","/屏蔽","/邀请舞伴","/对赌银两","/门派大战","/宠物","/如沐春风","/乾坤一掷","/拐骗少女","/拐骗少男","/同归于尽") ; 
if(lingualnum<maxlingual){lingualarr[lingualnum]=this.f2.document.af.sytemp.value;lingualnum++;}
else
{for (i=0;i<maxlingual;i++){lingualarr[i]=lingualarr[i+1];}lingualarr[i]=this.f2.document.af.sytemp.value;}
var pos=lingualnum;
var cmd=msg.substring(0,msg.length>1?1:msg.length);
var cmd1=msg.substring(0,msg.length>2?2:msg.length);
var saystemp;
saystemp=this.f2.document.af.sytemp.value
//while(saystemp.indexOf(" ") != -1 ){
//  saystemp=saystemp.replace(" ","")
//  }
var aftowho = this.f2.document.af.towho.value
if(IsBadWord(saystemp)){Warning();return false;}
for(i=0;i<sjcz.length;i++)
if(saystemp.search(sjcz[i]) != -1){ 
if(aftowho== "大家" ||aftowho == myn || aftowho==automan ) {
alert("不能对大家、机器人或者自己使用本功能！"+sjcz[i]+aftowho+automan);
this.f2.document.af.sytemp.value="";
return false;
}}
if(cmd=="/"& cmd1!="//")
{var spacepoint=msg.indexOf("$");
var spacepoint=(spacepoint==-1?msg.length:spacepoint);
var cmd=msg.substring(1,spacepoint);
switch(cmd){
case "主题":act="sjfunc/1.asp";break;case "振臂一呼":act="sjfunc/900.asp";break;
case "解穴":act="sjfunc/3.asp";break;case "逮捕":act="sjfunc/4.asp";break; 
case "坐牢":act="sjfunc/5.asp";break;case "监禁":act="sjfunc/6.asp";break; 
case "释放":act="sjfunc/7.asp";break;case "斩首":act="sjfunc/8.asp";break; 
case "警告":act="sjfunc/9.asp";break;case "下毒":act="sjfunc/10.asp";break; 
case "偷钱":act="sjfunc/11.asp";break;case "吸星大法":act="sjfunc/12.asp";break;  
case "投掷":act="sjfunc/13.asp";break;case "发招":act="sjfunc/14.asp";break; 
case "传授":act="sjfunc/15.asp";break;case "送钱":act="sjfunc/16.asp";break; 
case "赠送":act="sjfunc/17.asp";break;case "状态":act="sjfunc/18.asp";break; 
case "拍卖":act="sjfunc/19.asp";break;case "册封":act="sjfunc/20.asp";break; 
case "公告":act="sjfunc/21.asp";break;case "踢人":act="sjfunc/22.asp";break; 
case "心动":act="sjfunc/23.asp";break;case "放大":act="sjfunc/24.asp";break; 
case "拜师":act="sjfunc/25.asp";break;case "收徒":act="sjfunc/26.asp";break; 
case "打坐":act="sjfunc/27.asp";break;case "闭目":act="sjfunc/28.asp";break; 
case "经验":act="sjfunc/29.asp";break;case "练武":act="sjfunc/30.asp";break; 
case "存钱":act="sjfunc/31.asp";break;case "取钱":act="sjfunc/32.asp";break; 
case "转账":act="sjfunc/33.asp";break;case "清除":act="sjfunc/34.asp";break; 
case "修练":act="sjfunc/35.asp";break;case "暴豆":act="sjfunc/36.asp";break; 
case "哑穴":act="sjfunc/37.asp";break;case "好友":act="sjfunc/38.asp";break; 
case "怒吼":act="sjfunc/39.asp";break;case "心跳":act="sjfunc/40.asp";break; 
case "查ip":act="sjfunc/41.asp";break;case "奖励":act="sjfunc/42.asp";break; 
case "基金":act="sjfunc/43.asp";break;case "用卡":act="sjfunc/44.asp";break; 
case "招收弟子":act="sjfunc/46.asp";break; case "轩辕":act="sjfunc/14xx.asp";break;
case "求婚":act="sjfunc/47.asp";break;case "幸运":act="sjfunc/48.asp";break; 
case "赌博":act="sjfunc/49.asp";break;case "下注":act="sjfunc/50.asp";break; 
case "对赌":act="sjfunc/51.asp";break;case "清屏":act="sjfunc/52.asp";break; 
case "情人":act="sjfunc/53.asp";break;case "分手":act="sjfunc/54.asp";break; 
case "单挑":act="sjfunc/55.asp";break;case "炸弹":act="sjfunc/56.asp";break; 
case "送花":act="sjfunc/57.asp";break;case "本派刑法":act="sjfunc/58.asp";break; 
case "本派罚款":act="sjfunc/59.asp";break;case "掌门令":act="sjfunc/60.asp";break; 
case "赌豆":act="sjfunc/61.asp";break;case "暂离":act="sjfunc/62.asp";break; 
case "回来":act="sjfunc/63.asp";break;case "离席":act="sjfunc/64.asp";break; 
case "双人赌博":act="sjfunc/65.asp";break;case "使用":act="sjfunc/66.asp";break; 
case "丢弃":act="sjfunc/67.asp";break;case "监禁ip":act="sjfunc/68.asp";break; 
case "封锁ip":act="sjfunc/69.asp";break;case "逐出":act="sjfunc/70.asp";break; 
case "标题":act="sjfunc/71.asp";break;case "配药":act="sjfunc/72.asp";break; 
case "宠物":act="sjfunc/73.asp";break;case "竞标":act="sjfunc/74.asp";break; 
case "通缉":act="sjfunc/75.asp";break;case "解除":act="sjfunc/75.asp";break; 
case "出家":act="sjfunc/76.asp";break;case "还俗":act="sjfunc/76.asp";break; 
case "休身养性":act="sjfunc/77.asp";break;case "唸经":act="sjfunc/78.asp";break; 
case "如沐春风":act="sjfunc/78.asp";break;case "送豆点":act="sjfunc/79.asp";break; 
case "送金币":act="sjfunc/80.asp";break;case "乾坤一掷":act="sjfunc/81.asp";break; 
case "加入天网":act="sjfunc/82.asp";break;case "离开天网":act="sjfunc/82.asp";break; 
case "拐骗少女":act="sjfunc/83.asp";break;case "拐骗少男":act="sjfunc/84.asp";break; 
case "传授武功":act="sjfunc/87.asp";break;case "传授体力":act="sjfunc/88.asp";break;
case "宠物自暴":act="sjfunc/90.asp";break;case "寻找金币":act="sjfunc/91.asp";break;
case "同归于尽":act="sjfunc/92.asp";break;case "偷金币":act="sjfunc/93.asp";break;
case "送银币":act="sjfunc/94.asp";break;case "寻找银币":act="sjfunc/95.asp";break;
case "赌银币":act="sjfunc/96.asp";case "乞讨神术":act="sjfunc/45001.asp";break;
case "传送法力":act="sjfunc/45002.asp";break;case "存法力":act="sjfunc/450020.asp";break;
case "取法力":act="sjfunc/450021.asp";break;case "治病":act="sjfunc/501.asp";break;
case "教武":act="sjfunc/502.asp";break;case "求签":act="sjfunc/503.asp";break;
case "发射":act="sjfunc/601.asp";break;case "百变术":act="sjfunc/45007.asp";break;
case "布施术":act="sjfunc/45008.asp";break;case "魔界咒语":act="sjfunc/45009.asp";break;
case "魅惑人间":act="sjfunc/45011.asp";break;case "破天锥":act="sjfunc/45013.asp";break;
case "迷魂大法":act="sjfunc/450010.asp";break;case "爱神":act="sjfunc/85.asp";break;
case "颠倒":act="sjfunc/86.asp";break;case "移动":act="npc/a1.asp";break;
case "寻水晶球":act="sjfunc/45003.asp";break;case "魔幻水晶":act="sjfunc/45004.asp";break;
case "寻找法器":act="sjfunc/45005.asp";break;case "执行":act="sjfunc/602.asp";break;
case "没收法器":act="sjfunc/450151.asp";break;case "配制令牌":act="sjfunc/45015.asp";break;
case "天堂令":act="sjfunc/45012.asp";break;case "九阳神功":act="sjfunc/45013.asp";break;
case "使出":act="sjfunc/45019.asp";break;case "传授轻功":act="npc/1.asp";break;
case "轻功暂存":act="npc/2.asp";break;case "提出轻功":act="npc/3.asp";break;
case "讨取轻功":act="npc/4.asp";break;case "寻找秘笈":act="npc/5.asp";break;
case "宝物术":act="sjfunc/2001.asp";break;case "点金术":act="sjfunc/2003.asp";break;
case "解毒术":act="sjfunc/2004.asp";break;case "平安术":act="sjfunc/2005.asp";break;
case "偷窃术":act="sjfunc/2006.asp";break;case "摇钱术":act="sjfunc/2008.asp";break;
case "蓝玛瑙":act="sjfunc/2009.asp";break;case "帅哥令":act="sjfunc/2010.asp";break;
case "美人令":act="sjfunc/2011.asp";break;case "多情环":act="sjfunc/2012.asp";break;
case "绝命钩":act="sjfunc/2014.asp";break;case "暂存智力":act="sjfunc/1001.asp";break;
case "提取智力":act="sjfunc/1002.asp";break;case "粗体字":act="sjfunc/1003.asp";break;
case "飞舞字":act="sjfunc/1004.asp";break;case "按钮字":act="sjfunc/1005.asp";break;
case "滚动按钮":act="sjfunc/1006.asp";break;case "上下按钮":act="sjfunc/1007.asp";break; 
case "说明":act="ico/02.asp";break;case "倒夜香":act="sjfunc/505.asp";break;
case "赌法力":act="sjfunc/1012.asp";break;case "赌轻功":act="sjfunc/1013.asp";break;
case "邀请跳舞":act="sjfunc/1014.asp";break;case "查看公告":act="ico/09.asp";break;
case "小孩":act="sjfunc/98.asp";break;case "生育":act="sjfunc/97.asp";break;
case "原子弹":act="sjfunc/100.asp";break;case "练金属性":act="sjfunc/101.asp";break;
case "练木属性":act="sjfunc/102.asp";break;case "练水属性":act="sjfunc/103.asp";break;
case "练火属性":act="sjfunc/104.asp";break;case "练土属性":act="sjfunc/105.asp";break;
case "送金属性":act="sjfunc/106.asp";break;case "送木属性":act="sjfunc/107.asp";break;
case "送水属性":act="sjfunc/108.asp";break;case "送火属性":act="sjfunc/109.asp";break;
case "送土属性":act="sjfunc/110.asp";break;case "赌金属性":act="sjfunc/111.asp";break;
case "赌木属性":act="sjfunc/112.asp";break;case "赌水属性":act="sjfunc/113.asp";break;
case "赌火属性":act="sjfunc/114.asp";break;case "赌土属性":act="sjfunc/115.asp";break;
case "新人费":act="sjfunc/116.asp";break;case "修炼法力":act="sjfunc/117.asp";break;
case "修炼轻功":act="sjfunc/118.asp";break;case "传授道德":act="sjfunc/119.asp";break;
case "传授魅力":act="sjfunc/120.asp";break;case "二号情人":act="sjfunc/202.asp";break;
case "三号情人":act="sjfunc/203.asp";break;case "二号分手":act="sjfunc/206.asp";break;
case "三号分手":act="sjfunc/207.asp";break;case "超级令":act="sjfunc/99.asp";break;
case "打扑克":act="f2/dpk-ask.asp";break;case "发牌":act="f2/dpkfp.asp";break;
case "打麻将":act="f2/dmj-ask.asp";break;case "出牌":act="f2/DMJFP.ASP?action=1";break;
case "摸牌":act="f2/DMJFP.ASP?action=2";break;case "问牌":act="f2/DMJFP.ASP?action=3";break;
case "吃牌":act="f2/DMJFP.ASP?action=4";break;case "碰牌":act="f2/DMJFP.ASP?action=5";break;
case "杠牌":act="f2/DMJFP.ASP?action=6";break;case "招收官府":act="sjfunc/5042.asp";break;
case "偷金币":act="sjfunc/93.asp";break;case "对赌银两":act="sjfunc/zjn_add_fun_doub.asp";break;
case "邀请舞伴":act="sjfunc/tw.asp";break;case "门派大战":act="empdz/mp1.asp";break;
case "夺宝大战":act="sjfunc/db14.asp";break;case "夺宝下毒":act="sjfunc/db10.asp";break;
case "夺宝投掷":act="sjfunc/db13.asp";break;case "夺宝小孩攻击":act="sjfunc/dbxhgj.asp";break;
case "宠物夺宝":act="sjfunc/db73.asp";break;case "夺宝宠物自爆":act="sjfunc/db85.asp";break;
case "夺宝胜利":act="sjfunc/dbsl.asp";break;case "紫金葫芦":act="sjfunc/db35.asp";break;
case "屏蔽":act="sjfunc/02.asp";break;case "发射子弹":act="sjfunc/450017.asp";break;
case "回魂口诀":act="sjfunc/hhkj.asp";break;case "妙手回春":act="sjfunc/mshc.asp";break;
case "招财进宝":act="sjfunc/zcjb.asp";break;case "养心大法":act="sjfunc/yxdf.asp";break;
case "点石成金":act="sjfunc/dscj.asp";break;case "新雁南飞":act="sjfunc/xynf.asp";break;
case "新阴那山":act="sjfunc/xyns.asp";break;case "夺宝宠物攻击":act="sjfunc/db73.asp";break;
case "门派挑战":act="empdz/mpopen.asp";break;case "门派求和":act="empdz/mpopen1.asp";break;
case "血滴子":act="sjfunc/45014.asp";break;case "绝情刀":act="sjfunc/450019.ASP";break;
case "抢劫令":act="sjfunc/45016.asp";break;case "没收魔器":act="sjfunc/450152.asp";break;
case "寻找魔器":act="sjfunc/45017.asp";break;case "配制宝石":act="sjfunc/45018.asp";break;
case "生日蛋糕":act="sjfunc/450018.ASP";break;case "字体魔法":act="sjfunc/1008.asp";break;
case "移动魔法":act="sjfunc/1008.asp";break;case "按钮魔法":act="sjfunc/1008.asp";break;
case "魔力钻石":act="sjfunc/45006.asp";break;case "百变神通":act="sjfunc/45007.asp";break;
case "夺宝轩辕":act="sjfunc/14dbxx.asp";break;case "邀请赏花":act="sjfunc/yqsh.asp";break;
case "轰炸窗口":act="sjfunc/ckzd.asp";break;case "轰炸硬盘":act="sjfunc/gpzd.asp";break;case "今日话题":act="sjfunc/45.asp";break;
default:alert('不明命令,无法执行')
this.f2.document.af.oldtowho.value=this.f2.document.af.towho.value;
this.f2.document.af.oldtowho.value=this.f2.document.af.towho.value;
this.f2.document.af.sy.value=saystemp;
this.f2.document.af.oldsays.value=saystemp;
this.f2.document.af.addsign.options[0].selected=true;
this.f2.document.af.tu.options[0].selected=true;
this.f2.document.af.sytemp.focus();
this.f2.document.af.sytemp.value='';return false;}}
else
if (nowmsg=="//"){act="say.asp";}else{act="say.asp"}
this.f2.document.af.action=act;
//this.f2.document.af.sjjhaction.value=act;
{this.f2.document.af.sy.value='';if(saystemp!=''){if((this.f2.document.af.oldsays.value==saystemp)&&(this.f2.document.af.oldtowho.value==this.f2.document.af.towho.value)){alert('内容不可重复！');
this.f2.document.af.sytemp.focus();this.f2.document.af.sytemp.select();return false;}this.f2.document.af.oldtowho.value=this.f2.document.af.towho.value;
this.f2.document.af.sy.value=saystemp;this.f2.document.af.oldsays.value=saystemp;this.f2.document.af.addsign.options[0].selected=true;
this.f2.document.af.tu.options[0].selected=true;this.f2.document.af.sytemp.focus();this.f2.document.af.sytemp.value='';
ty=new Date();var nh=ty.getHours();var nm=ty.getMinutes();var ns=ty.getSeconds();var ct=(nh*3600)+(nm*60)+ns;
if(((ct-lst)<2.5)&&(ct>lst)){this.f2.af.sytemp.value=this.f2.af.oldsays.value;this.f2.af.oldsays.value='';return false;}
else{lst=ct;}this.f2.addOne(this.f2.document.af.sy.value);this.f2.startnosay();
this.f2.document.af.subsay.disabled=1;setTimeout('this.f2.document.af.subsay.disabled=0',3000);return true;}if ((this.f2.document.af.addsign.options[this.f2.document.af.addsign.selectedIndex].value=='0')||(this.f2.document.af.addsign.options[this.f2.document.af.addsign.selectedIndex].value=='')){alert('请输入发言或选择动作！');
this.f2.document.af.sytemp.focus();this.f2.document.af.sytemp.select();return false;}}}
function wdmess(b){this.mess.document.writeln(b);}
function mytitle(title){
this.title.document.open();
this.title.document.writeln("<html><head><meta http-equiv=\"content-type\" content=\"text/html; charset=gb2312\"><\/head>");
this.title.document.writeln("<style type=\"text/css\">body{font-family:\"宋体\";color:blue;font-size:10.5pt;line-height:15pt;text-align:center}</style>");
this.title.document.writeln("<body bgcolor='#b7d4f1' oncontextmenu=window.event.returnValue=false onselectstart=event.returnValue=false ondragstart=window.event.returnValue=false TOPMARGIN=2 LEFTMARGIN=5 MARGINWIDTH=0 MARGINHEIGHT=0><div align='center'>");
this.title.document.writeln(title+"</div></body></html>");
this.title.document.close();}
function DB(){this.mess.document.open();
wdmess("<html><head><meta http-equiv=\"content-type\" content=\"text/html; charset=gb2312\"><\/head>");
wdmess("<style type=\"text/css\">body{font-family:\"宋体\";color:blue;font-size:10.5pt;line-height:15pt;text-align:center}a{color:blue;text-decoration:none;}a:hover{color:blue;text-decoration:underline;}</style>");
wdmess("<body bgcolor='#eaeaff' TOPMARGIN=2 LEFTMARGIN=5 MARGINWIDTH=0 MARGINHEIGHT=0>");
var i=0;i=Math.ceil(Math.random()*(Banner.length-1));
wdmess("<marquee scrollamount=3 scrolldelay=10 onmouseover=this.stop(); onmouseout=this.start();>" + Banner[i] + "<\/marquee><\Script Language=JavaScript>setTimeout('parent.DB()',100000);<\/script></body></html>");
this.mess.document.close();}

function tbclutch(){if(this.f2.document.af.tbclutch.value=='全屏'){this.f2.document.af.tbclutch.value='垂直';this.msgfrm.rows = "*";this.msgfrm.cols = "*";tbclu=false;}else{if(this.f2.document.af.tbclutch.value=='垂直'){this.f2.document.af.tbclutch.value='水平';this.msgfrm.cols = "*,*";this.msgfrm.rows = "*";tbclu=true;}else{this.f2.document.af.tbclutch.value='全屏';this.msgfrm.cols = "*";this.msgfrm.rows = "*,*";tbclu=true;}}this.f2.document.af.sytemp.focus();}
function tbmd(){if(this.f2.document.af.tbmd.value=='→'){this.f2.document.af.tbmd.value='←';this.tbgn1.cols="1*,0";}else{this.f2.document.af.tbmd.value='→';this.tbgn1.cols="1*,160";}this.f2.document.af.sytemp.focus();}
function tbygn(){if(this.f2.document.af.tbygn.value=='↓'){this.f2.document.af.tbygn.value='↑';this.tbymd.rows="0,*,0,0" ;}else{this.f2.document.af.tbygn.value='↓';this.tbymd.rows="0,*,0,105" ;}this.f2.document.af.sytemp.focus();}
function tbgn(){if(this.f2.document.af.tbgn.value=='功能开'){this.f2.document.af.tbgn.value='功能关';this.mainfrm.cols="34,*";}else{this.f2.document.af.tbgn.value='功能开';this.mainfrm.cols="0,*";}this.f2.document.af.sytemp.focus();}
self.onerror=null;
var nullframe = '<HTML><BODY BGCOLOR=#000000 text=#ffffff><center><H3 color=yellow><br><font color=yellow>欢迎来到≮<%=Application("sjjh_chatroomname")%>总站≯</font></h3><br><font size=4>程序正在从服务器读取资料, 请稍候 ......</font></center></BODY></HTML>';
</script>
</head>
<frameset cols="1,*,1,160,1" name=tbgn1 rows="*" border="0" framespacing="0" frameborder="NO">
	<frameset rows="47,*,72" cols="*" frameborder="no" name=tbymd7>
		<frame src="ico/wfy24.htm" name="llmbt1" scrolling="no">
		<frame src="ico/wfy26.htm" name="llmbt2" scrolling="no">
		<frame src="ico/wfy25.htm" name="llmbt3" scrolling="no">
	</frameset>

	<frameset rows="24,*,0,0,0,25,76,0,1" cols="*">
			<frame src="ico/eline_time.htm" name="yt" scrolling="no" marginwidth="0" marginheight="0">
		<frameset name=msgfrm rows="*,*" cols="*">
			<frame src="javascript:parent.nullframe" name="f0" scrolling="AUTO" framespacing="0" marginheight="3" marginwidth="5" frameborder="yes">
			<frame src="javascript:parent.nullframe" name="f1" scrolling="AUTO" framespacing="0" marginheight="3" marginwidth="5" frameborder="no">
		</frameset>

		<frame src="guanggao.asp" scrolling="NO"  name="gg" marginwidth="3" marginheight="3">
		<frame src="about:blank" name="mess" scrolling="no" >
		<frame src="about:blank" name="t" marginwidth="5" marginheight="5"  scrolling="NO">
		<frame src="about:blank" name="title" scrolling="no" >
		<frame src="f22.asp" name="f2" scrolling="NO" marginwidth="3" marginheight="8">
		<frame src="about:blank" name="d" scrolling="NO">
		<frame src="ico/wfy15.htm" name="zt" scrolling="no" marginwidth="0" marginheight="0">
	</frameset>

	<frameset rows="47,*,72" cols="*" frameborder="no" name=tbymd>
		<frame src="ico/wfy20.htm" name="lmbt1" scrolling="no">
		<frame src="ico/wfy19.htm" name="lmbt2" scrolling="no">
		<frame src="ico/wfy18.htm" name="lmbt3" scrolling="no">
	</frameset>
	<frameset rows="1,0,*,0,182" name=tbymd>
		<frame src="ico/wfy16.htm" name="st" scrolling="no">
			<frame src="about:blank" name="m">
			<frame src="about:blank" marginwidth="5" marginheight="5" name="f3">
			<frame src="about:blank" marginwidth="5" marginheight="5" name="ps">
			<frame src="F4.asp" name="F4" scrolling="no">
	</frameset>
	<frameset rows="47,*,72" cols="*" frameborder="no" name=tbymd>
		<frame src="ico/wfy23.htm" name="rmbt1" scrolling="no">
		<frame src="ico/wfy22.htm" name="rmbt2" scrolling="no">
		<frame src="ico/wfy21.htm" name="rmbt3" scrolling="no">
	</frameset>
</frameset>

</noframes>
</html>