<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../config.asp"-->
<!--#include file="../mywp.asp"-->
<%Response.Buffer=true
Response.Expires = 0
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
chatroomsn=session("nowinroom")	'当前所在房间序号
mychatroomsn=Request.QueryString("roomsn")	'要去的房间序号
chatroomname=Application("sjjh_chatroomname"&mychatroomsn)
mm=Application("sjjh_chatroomname"&chatroomsn)
if chatroomsn=mychatroomsn then
  response.write "<Script language=javascript>alert('你已经在『"&chatroomname&"』不能重复进入！');parent.m.location.reload();</script>"
  response.end
end if
if (sjz>72020 and sjz<72050) and mm="高手房间" and sjjh_grade<6 then
	response.write "<Script language=javascript>alert('即然已经走进『"&chatroomname&"』参赛，则无路可退，除非你退出聊天室！');parent.m.location.reload();</script>"
	response.end
end if
	n=Year(date())
	y=Month(date())
	r=Day(date())
	s=Hour(time())
	f=Minute(time())
	m=Second(time())
	weekdate=weekday(date())
	sjz=weekdate*10000+s*100+f
	'星期五晚20:00则为6*10000+20*100+0=62000
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
t=s & ":" & f & ":" & m
sj=n & "-" & y & "-" & r & " " & t

chatroominfo=split(Application("sjjh_room"),";")
chatroomnum=ubound(chatroominfo)-1
i=mychatroomsn
online=split(trim(Application("sjjh_useronlinename"&i))," ")	'取房间人员名单
onlinenum=ubound(online)+1
sj_chat_info=split(chatroominfo(i),"|")	'分离要去的房间数据
if sj_chat_info(0)="高手房间" and sjjh_grade<>10 then
	if (sjz>71200 and sjz<72020) or sjz>72030 then
		erase sj_chat_info
		erase online
		erase chatroominfo
		response.write "<Script language=javascript>alert('今天乃夺宝之日，只能在晚上20:20-20:30之间进入！');parent.m.location.reload();</script>"
		response.end
	else
		if sjz>71200 and sjz<72020 then
			erase chatroominfo
			erase sj_chat_info
			erase online
			response.write "<Script language=javascript>alert('今天乃夺宝之日，进入时间为晚上20:20至20:30，宝物夺取开始时间为20:30整。');parent.m.location.reload();</script>"
			response.end
		else
			if sjz>72030 then
				erase sj_chat_info
				erase online
				erase chatroominfo
				response.write "<Script language=javascript>alert('已过了夺宝进入时间，该时间为每星期六晚20:20至20:30，夺宝大赛开始时间为20:00整，下次记着早点进入哟！');parent.m.location.reload();</script>"
				response.end
			end if
		end if
	end if
end if
if (sjz>=72020 and sjz<=72030) and (sjjh_grade>=6 and sjjh_name<>"回首当年") and chatroomname="高手房间" then
		Response.Write "<script language=JavaScript>{alert('提示：官府人员不可参与夺宝！');}</script>"
		Response.End
	end if

num=int(sj_chat_info(1))
if onlinenum>=num then
	erase chatroominfo
	erase sj_chat_info
	erase online
	response.write "<Script>alert('『"&chatroomname&"』房间当前人数已满，不能进入！');;</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
if cstr(sj_chat_info(2))=1 then
	if sjjh_grade<7 then
		rs.open "select id,姓名,等级,身份,门派,体力加,内力加,名单头像,性别,会员等级,操作时间,w1 from 用户 where 姓名='" & sjjh_name &"'"&" and "&sj_chat_info(4),conn,2,2
	else
		rs.open "select id,姓名,等级,身份,门派,体力加,内力加,名单头像,性别,会员等级,操作时间,w1 from 用户 where 姓名='" & sjjh_name &"'",conn,2,2
	end if
else
	rs.open "select id,姓名,等级,身份,门派,体力加,内力加,名单头像,性别,会员等级,操作时间,w1 from 用户 where 姓名='" & sjjh_name &"'",conn,2,2
end if
if cstr(sj_chat_info(2))=1 and sjjh_grade<7 then
	if rs.eof or rs.bof then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<html><head><meta http-equiv='Content-Type' content='text/html; charset=gb2312'><meta http-equiv='pragma' content='no-cache'></head><body bgcolor=" + Application("sjjh_chatbgcolor") + " background=" + Application("sjjh_chatimage") + " bgproperties=fixed>"
		Response.Write "<script language=JavaScript>{alert('进入『"&sj_chat_info(0)&"』的条件是："&sj_chat_info(3)&"');}</script>"
		erase chatroominfo
		erase sj_chat_info
		erase online
		Response.End
	end if
end if
sjjh_id=rs("id")
jhsf=rs("身份")
jhmp=rs("门派")
jhtx=rs("名单头像")
sex=rs("性别")
hydj=rs("会员等级")
sj=DateDiff("n",rs("操作时间"),now())
tlsx=rs("等级")*sjjh_tlsx+5260+rs("体力加")
nlsx=rs("等级")*sjjh_nlsx+2000+rs("内力加")
w1=rs("w1")
rs.close
if sj<3 and sjjh_grade<6 and sj_chat_info(0)<>"高手房间" then
	s=3-sj
	set rs=nothing	
	conn.close
	set conn=nothing
	erase chatroominfo
	erase sj_chat_info
	erase online
	Response.Write "<script language=JavaScript>{alert('提示：转换房间请等["&s&"分钟]再操作！');parent.m.location.reload();}</script>"
	Response.End
end if
if sjz>=72020 and sjz<=72030 and jhmp="出家" and chatroomname="高手房间" then
        Response.Write "<script language=JavaScript>{alert('提示：出家人不可参与夺宝！');parent.m.location.reload();}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if (sjz>72020 and sjz<72030) and sjjh_grade<6 and sj_chat_info(0)="高手房间" then
	a=mywpsl(w1,"千年人参")
	b=mywpsl(w1,"万年灵芝")
	butl=a*5000
	bunl=b*5000
	tlyj=int(tlsx/5000)
	nlyj=int(nlsx/5000)
	if butl>tlsx or bunl>nlsx then
		set rs=nothing	
		conn.close
		set conn=nothing
		erase chatroominfo
		erase sj_chat_info
		erase online
		Response.Write "<script language=JavaScript>{alert('参加夺宝大赛时只允许带补足一次体力内力的千年人参及万年灵芝，你的药太多，只能带["&tlyj&"]个千年人参，["&nlyj&"]个万年灵芝。');parent.m.location.reload();}</script>"
		Response.End
	end if
end if
'姓名,性别,门派,身份,头象,等级,id
myzanli=0
if Instr(LCase(application("sjjh_zanli")),LCase("!"&sjjh_name&"!"))>0 then myzanli=1
if myzanli=1 and (sj_chat_info(0)="高手房间" or sj_chat_info(0)="快乐江湖") then
	set rs=nothing	
	conn.close
	set conn=nothing
	erase chatroominfo
	erase sj_chat_info
	erase online
	Response.Write "<script language=JavaScript>{alert('提示：你在暂离中，不可以进入此房间！');parent.m.location.reload();}</script>"
	Response.End
end if
myonline = sjjh_name & "|" & sex & "|" & jhmp & "|" & jhsf & "|" & jhtx & "|" & sjjh_jhdj& "|" & sjjh_id& "|" & hydj&"|"&myzanli&"|"&"正常"
if sj_chat_info(7)<>0 then
	conn.Execute "update 用户 set 保护=false,操作时间=now() where 姓名='" & sjjh_name &"'"
else
	conn.Execute "update 用户 set 保护=true,操作时间=now() where 姓名='" & sjjh_name &"'"
end if
set rs=nothing	
conn.close
set conn=nothing
'退出原房间
Application.Lock
mychatroomname=Application("sjjh_chatroomname"&chatroomsn)
onlinelist=Application("sjjh_onlinelist"&chatroomsn)
onno=ubound(onlinelist)
for i=1 to onno
 if InStr(onlinelist(i),sjjh_name & "|")=1 then
	  for j=i to onno-1
		onlinelist(j)=onlinelist(j+1)
	  next
	  ReDim Preserve onlinelist(onno-1)
	  Application("sjjh_onlinelist"&chatroomsn)=onlinelist
	  exit for
 end if
next
Application("sjjh_useronlinename"&chatroomsn)=Replace(Application("sjjh_useronlinename"&chatroomsn)," " & sjjh_name & " ","")
'进入新房间
dim newonlinelist()
js=1
onlinelist=Application("sjjh_onlinelist"&mychatroomsn)
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
if yjl=0  then
	Redim Preserve newonlinelist(js)
	newonlinelist(js)=myonline
end if
Application("sjjh_onlinelist"&mychatroomsn)=newonlinelist
Application("sjjh_useronlinename"&mychatroomsn)=Application("sjjh_useronlinename"&mychatroomsn)& " "&sjjh_name & " "
erase  newonlinelist
erase  onlinelist
session("nowinroom")=mychatroomsn
Application.UnLock	

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
act="退出"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
if (sjz>72020 and sjz<72030) and sj_chat_info(0)="高手房间" then
	says="<font color=black>【换房】</font><font color=#009933>一周一次的夺宝时间终于到了，<font color=red>" & sjjh_name & "</font>满怀自信地从<font color=red>〖" & mychatroomname & "〗</font>阔步迈进了<font color=red>【" &chatroomname & "】</font>房间，让我们大家祝他好运吧！</font><font class=t>(" & time() & ")</font><bgsound src='readonly/cd.mid' loop='1'>"
else
	says="<font color=black>【换房】</font><font color=#009933><font color=red>" & sjjh_name & "</font>施展出“凌波微步”轻功，转眼间便从<font color=red>〖" & mychatroomname & "〗</font>消失了，原来是去<font color=red>【" &chatroomname & "】</font>了。</font><font class=t>(" & time() & ")</font><bgsound src='readonly/cd.mid' loop='1'>"
end if
says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
saystr="<"&"script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & sjjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &",0," & chatroomsn & ");<"&"/script>"
addmsg saystr
'session("SayCount")=Application("SayCount")
act="进入"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
if (sjz>72020 and sjz<72030) and sj_chat_info(0)="高手房间" then
	says="<font color=#cc0000>【上线】</font><font color=#009933>一周一次的夺宝时间终于到了，<a href=javascript:parent.sw('『" & sjjh_name & "』'); target=f2>" & sjjh_name & "</a>满怀自信地从<font color=red>〖" & mychatroomname & "〗</font>阔步迈进了<font color=red>【" &chatroomname & "】</font>房间，让我们大家祝他好运吧！</font><font class=t>(" & t & ")</font><bgsound src='readonly/cdcd.wav' loop='1'>"	'受话者
else
	says="<font color=#cc0000>【上线】</font><font color=#009933><a href=javascript:parent.sw('『" & sjjh_name & "』'); target=f2>" & sjjh_name & "</a>施展出“凌波微步”轻功，转眼间从〖" & mychatroomname & "〗来到了【" &chatroomname & "】。</font><font class=t>(" & t & ")</font><bgsound src='readonly/cdcd.wav' loop='1'>"	'受话者
end if

says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))

saystr="<"&"script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & sjjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &",0," & session("nowinroom") & ");<"&"/script>"
addmsg saystr
For i = 1 to Application("SayCount")-Session("SayCount")
	Response.Write Application("SayStr"&YuShu((Session("SayCount")+i)))
Next
session("SayCount")=Application("SayCount")
%>
<Script >
parent.crm='<%=Application("sjjh_chatroomname"&session("nowinroom"))%>';
parent.myroom=<%=session("nowinroom")%>
parent.m.location.reload();
//parent.location.href='jhchat.asp';
parent.f2.location.href='f22.asp?id=1';
alert('欢迎来到『<%=chatroomname%>』');
</script>