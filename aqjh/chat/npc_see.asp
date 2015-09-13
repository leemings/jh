<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../const3.asp"-->
<%Response.Buffer=true
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Response.Expires=0
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
aqjh_roominfo=split(Application("aqjh_room"),";")
chatroomnum=ubound(aqjh_roominfo)-1
for i=0 to chatroomnum	
	ydl=1
	if Instr(LCase(Application("aqjh_useronlinename"&i))," "&LCase(aqjh_name)&" ")=0 then ydl=0
	if ydl=1 and clng(nowinroom)<>i then 
		Session.Abandon
		Response.Redirect "../error.asp?id=140"
		Response.End 
	end if
        aqjh_roominfo=split(Application("aqjh_room"),";")
        chatinfo=split(aqjh_roominfo(nowinroom),"|")
        if chatinfo(0)<>aqjh_chat3 then
	Response.Write "<script language=JavaScript>{alert('提示:NPC只有在["&aqjh_chat3&"]内有效！');window.close();}</script>"
	Response.End
end if
if aqjh_ifnpc=0 then
        Response.Write "<script language=JavaScript>{alert('提示:NPC暂时被系统关闭！');window.close();}</script>"
	Response.End
end if
next 
'房间
chatroomname=trim(Application("aqjh_chatroomname"&session("nowinroom")))
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
'取当前在线npc总数
npcsl=ubound(split(Application("aqjh_npc"),";"))
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
sj2=n & "-" & y & "-" & r & " " & sj
id=clng(Request.QueryString("id"))
act=Request.QueryString("act")
Response.Write "<script>window.moveTo(100,30);document.onmousedown=click;}</script>"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
jbmoney=1
if act<>"" then
	'决战判断
	act=clng(act)
	if act<>1 and act<>2 and act<>3 then
		act=1
	end if
	if session("aqjh_grade")<application("aqjh_npcff") and act>1 then
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script Language=Javascript>alert('提示：" & aqjh_name & "你没有发放NPC的权力!');location.href = 'npc_see.asp?id=" & id & "';</script>"
		Response.end
	end if
	select case act
	case 1'唤醒npc使用的是金币这个与怪物等相相关，每一级怪物就是一个金币大家可以改jbmoney=1来自己改倍数
		'取出npc等级计算唤醒他需要的金币数
		if npcsl>=Application("aqjh_maxnpc") Then 
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script Language=Javascript>alert('提示：" & aqjh_name & "江湖规定江湖npc在线数量最多["& Application("aqjh_maxnpc") &"]个!');location.href = 'npc_see.asp?id=" & id & "';</script>"
			Response.end
		end if
		rs.open "Select  * from [npc] where id=" & id,conn,1,1
		If rs.bof  Then
			rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script Language=Javascript>alert('提示：所查找的NPC并不存在!');window.close();</script>"
			Response.end
		end if
		jbsl=(int(sqr(int(rs("n经验")/50)))*jbmoney)
		name=rs("n姓名")
		if InStr(";" & Application("aqjh_npc"), ";" & name & "|")<>0 then
			rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script Language=Javascript>alert('提示：此NPC现在已经在线了，你不能进行唤醒操作!');location.href = 'npc_see.asp?id=" & id & "';</script>"
			Response.end
		end if
		zuren=0
		says="<font color=#ff0000>【NPC消息】</font>["&aqjh_name&"]在恶魔谷找到了<a href=javascript:parent.sw(\'[" & name & "]\'); target=f2>" & name &"</a>，用金钱将"&name&"收买作了自己的小弟,有了"&name&"的帮助["&aqjh_name&"]实力大增!</font>！"
		if rs("n主人")=session("aqjh_name") then
			jbsl=int(jbsl/2+1)
			says="<font color=#ff0000>【NPC消息】</font>["&aqjh_name&"]在恶魔谷终于找到了自己的老步下<a href=javascript:parent.sw(\'[" &name & "]\'); target=f2>" & name &"</a>，有了"&name&"的帮助["&aqjh_name&"]实力大增!</font>！"
		end if
		rs.close
		'判断自己的金币数量,如果这个npc是我的，我花一半钱召出来
			rs.open "Select * from [用户] where 姓名='"& session("aqjh_name") &"'",conn,1,3
			if rs("金币")<jbsl then
				rs.close
				set rs=nothing
				conn.close
				set conn=nothing
				Response.Write "<script Language=Javascript>alert('提示：你没有足够的金和币，不能唤醒此NPC!');location.href = 'npc_see.asp?id=" & id & "';</script>"
				Response.end
			end if
                        if rs("职业")<>"魔法师" then
				rs.close
				set rs=nothing
				conn.close
				set conn=nothing
				Response.Write "<script Language=Javascript>alert('提示：只有魔法师才能唤醒此NPC!');location.href = 'npc_see.asp?id=" & id & "';</script>"
				Response.end
			end if
			rs("金币")=rs("金币")-jbsl
			rs.update
			rs.close
		rs.open "Select * from [npc] where id=" & id,conn,1,3
		rs("n主人")=session("aqjh_name")
                id=rs("id")
		name=rs("n姓名")
                sex=rs("n性别")
                jhtx=rs("n图片")
                j=int(sqr(int(rs("n经验")/50)))
                rs("n攻击")=j*100
                rs("n防御")=j*150
                rs("n武功")=j*800
                rs("n体力")=j*5000
                rs("n内力")=j*150
                rs("n银两")=j*150000
                rs("n出现率")=rs("n出现率")+1
                rs("n复活时间")=now()
                rs("n攻击时间")=now()
                myonline = name & "|" & sex & "|NPC|"&rs("n主人")&"|" & jhtx & "|" &j& "|"&id&"|0|1|"&"正常"
		rs("n敌人")="无"
		rs.update
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
'招唤成功后添加名单
Application.Lock
js=1
onlinelist=Application("aqjh_onlinelist"&nowinroom)
onlineno=ubound(onlinelist)
yjl=0
for i=1 to onlineno
onuser=split(onlinelist(i),"|")
if yjl=0 and StrComp(onuser(0),aqjh_name,1)=1 then	'按名字汉语拼音排序
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
Application("aqjh_onlinelist"&nowinroom)=newonlinelist
Application("aqjh_useronlinename"&nowinroom)=Application("aqjh_useronlinename"&nowinroom)&" "&name&" "
Application("aqjh_npc")=Application("aqjh_npc")&";"&name&"|"
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
'招唤成功后添加名单
call showchat(says)
		Response.Write "<script Language=Javascript>alert('提示：成功招出NPC，现在你是他的主人，在与别\n人参战时他会帮助你的!');location.href = 'npc_see.asp?id=" & id & "';</script>"
		Response.end
	case 2,3
		rs.open "Select  * from [npc] where id=" & id,conn,2,3
		If rs.bof  Then
			rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script Language=Javascript>alert('提示：所查找的NPC并不存在!');window.close();</script>"
			Response.end
		end if
                id=rs("id")
		name=rs("n姓名")
                sex=rs("n性别")
                jhtx=rs("n图片")
                nccc=rs("n出场词")
                j=int(sqr(int(rs("n经验")/50)))
                rs("n攻击")=j*100
                rs("n防御")=j*150
                rs("n武功")=j*800
                rs("n体力")=j*5000
                rs("n内力")=j*150
                rs("n银两")=j*150000
                if act=2 then rs("n出现率")=rs("n出现率")+1
                rs("n敌人")="无"
                rs("n主人")="无"
                rs("n复活时间")=now()
                rs("n攻击时间")=now()
		rs.update
                myonline = name & "|" & sex & "|NPC|"&rs("n主人")&"|" & jhtx & "|" &j& "|"&id&"|0|1|"&"正常"
		rs.close
		if act=2 then
			if InStr(";" & Application("aqjh_npc"), ";" & name & "|")<>0 then
			Response.Write "<script Language=Javascript>alert('提示：此NPC已经在线了，不能再放出!');location.href = 'npc_see.asp?id=" & id & "';</script>"
			Response.end
			end if
                        npc_lasttime=DateDiff("n",Application("aqjh_npc_lasttime"),now())
                        if Application("aqjh_npc_lasttime")<>"" and npc_lasttime<90 and aqjh_name<>Application("aqjh_user") then
			Response.Write "<script Language=Javascript>alert('提示：官府及掌门元老每1个半小时发放一次！站长不受限制！');location.href = 'npc_see.asp?id=" & id & "';</script>"
			Response.end
                        end if
                        Application("aqjh_npc_lasttime")=now()
           says = "<bgsound src=readonly/cd.mid loop=1><font color=black>【NPC进入公告】<a href=javascript:parent.sw(\'[" & name & "]\'); target=f2>" & name &"</a></font>"&nccc
          show="NPC发放成功!"
'添加名单
Application.Lock
js=1
onlinelist=Application("aqjh_onlinelist"&nowinroom)
onlineno=ubound(onlinelist)
yjl=0
for i=1 to onlineno
onuser=split(onlinelist(i),"|")
if yjl=0 and StrComp(onuser(0),aqjh_name,1)=1 then	'按名字汉语拼音排序
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
Application("aqjh_onlinelist"&nowinroom)=newonlinelist
Application("aqjh_useronlinename"&nowinroom)=Application("aqjh_useronlinename"&nowinroom)&" "&name&" "
Application("aqjh_npc")=Application("aqjh_npc")&";"&name&"|"
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
'添加名单
call showchat(says)
		else 
                       if InStr(";" & Application("aqjh_npc"), ";" & name & "|")=0 then
				Response.Write "<script Language=Javascript>alert('提示：此NPC不在线你不能踢出他!');location.href = 'npc_see.asp?id=" & id & "';</script>"
				Response.end
			end if
'删除名单
nowinroom=session("nowinroom")
Application.Lock
onlinelist=Application("aqjh_onlinelist"&session("nowinroom"))
onno=ubound(onlinelist)
for i=1 to onno
if InStr(onlinelist(i),name & "|")=1 then
  for j=i to onno-1
   onlinelist(j)=onlinelist(j+1)
  next
  ReDim Preserve onlinelist(onno-1)
  exit for
 end if
next
Application("aqjh_onlinelist"&nowinroom)=onlinelist
Application("aqjh_useronlinename"&nowinroom)=Replace(Application("aqjh_useronlinename"&nowinroom)," " & name & " ","")
Application("aqjh_npc")=Replace(Application("aqjh_npc"),";"&name&"|","")
Application.UnLock
            says = "<bgsound src=wav/gf.wav loop=1><font color=black>【踢出NPC】</font>管理员["&aqjh_name&"]对NPC人物["&name&"]进行了踢出操作，"&name&"退出了江湖!"
            show="NPC踢出成功!"
call showchat(says)
		end if
		Response.Write "<script Language=Javascript>alert('提示："& show &"');location.href = 'npc_see.asp?id=" & id & "';</script>"
		Response.end
	end select
end if
rs.open "Select  * from [npc] where id=" & id,conn,1,1
If rs.bof  Then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：所查找的NPC并不存在!');window.close();</script>"
	Response.end
end if
name=rs("n姓名")
%>
<html>
<head>
<title>查看NPC[<%=name%>]的详细状态</title>
<LINK href="jh.css" rel=stylesheet>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body background="bg.gif" text="#FFFF00" leftmargin="0" topmargin="4" oncontextmenu=window.event.returnValue=false>
<table border="1" width="99%" cellpadding="1" cellspacing="0" bordercolordark="#FFFFFF" bordercolorlight="#00215a" align="center">
<tr align="center">
<td colspan="4">查看NPC[<%=name%>]的详细状态
<div id="Layer1" style="position:absolute; width:87px; height:73px; z-index:1; left: 404px; top: 370px;"><img src="<%=rs("n图片")%>"></div>
</td>
</tr>
<tr>
<td width="25%" align="left">名称：<%=name%></td>
<td width="25%" align="left">性别：<%=rs("n性别")%></td>
<td width="25%" align="left">等级：<%=int(sqr(int(rs("n经验")/50)))%></td>
<td width="25%" align="left">死亡：<%=rs("n死次数")%>次</td>
</tr>
<tr>
<td align="left">攻击：<%=rs("n攻击")%></td>
<td align="left">防御：<%=rs("n防御")%></td>
<td align="left">武功：<%=rs("n武功")%></td>
<td align="left">体力：<%=rs("n体力")%></td>
</tr>
<tr>
<td align="left">内力：<%=rs("n内力")%></td>
<td align="left">银两：<%=rs("n银两")%></td>
<td align="left">出现率：<%=rs("n出现率")%></td>
<td align="left">攻击率：<%=rs("n攻击率")%></td>
</tr>
<tr>
<td align="left">活:<%=rs("n复活时间")%></td>
<td align="left">死:<%=rs("n死亡时间")%></td>
<td align="left">主人：<%=rs("n主人")%></td>
<td align="left"><%online=InStr(";" & Application("aqjh_npc"), ";" & name & "|")
If online<>0 Then%>此NPC在线!<%else%>此NPC不在线<%end if%></td>
</tr>
<tr>
<td colspan="4" align="center">
<input type="submit" name="Submit" onclick="javascript:<%If online<>0 Then%>alert('此NPC已经在线不能再唤醒!');<%else%>if(confirm('唤醒此NPC需要金币：<%=(int(sqr(int(rs("n经验")/50)))*jbmoney)%>个，如果你是它的主人只需花一半的钱 \n唤醒你就是他的主人，你确定进行唤醒操作吗？')){location.href='npc_see.asp?id=<%=id%>&act=1';}<%end if%>" value="唤醒NPC">&nbsp;&nbsp;&nbsp;
<%if session("aqjh_grade")>=application("aqjh_npcff") or Session("aqjh_name") = Application("aqjh_user") then%>
<input type="submit" name="Submit2" onclick="javascript:<%If online<>0 Then%>alert('此NPC已经在线不能再发放!');<%else%>location.href='npc_see.asp?id=<%=id%>&act=2';<%end if%>" value="发放NPC">&nbsp;&nbsp;&nbsp;
<input type="submit" name="Submit3" onclick="javascript:<%If online=0 Then%>alert('此NPC并不在线不能再踢出!');<%else%>location.href='npc_see.asp?id=<%=id%>&act=3';<%end if%>" value="踢出NPC">
<%end if%>
</td>
</tr>
<tr>
<td colspan="4" align="left">
<p><strong>1.NPC死后的物品计算。</strong><br>
<font color="#FFFFFF">在打死NPC后物品的暴出值，NPC死后会根据杀人者等级，NPC等级进行计算：(NPC等级/杀人者等级)*1000然后取整数，再用此值与aqjh_npcwp进行比较，如果此计算的值大于等于aqjh_npcwp的值则进行物品生产，否则退出终止程序，这个主要是为了处理高级人员打低级NPC的比例。如果通过了比较物品也不是%100暴出的，会根据NPC死的次数与当前物品值进行计算随机得出来
。 <br>
aqjh_npcwp现在的值是:<strong><font color="#00FF00"><%=application("aqjh_npcwp")%></font></strong>
比如我:70级NPC是65级.(65/70)x1000=928此数如果大于<strong><font color="#00FF00"><%=application("aqjh_npcwp")%></font></strong>,现在我打这个npc会有物品出来.小于则没有!<br>
</font> <strong>2.NPC自动攻击计算.</strong><font color="#FFFFFF"><br>
        NPC会随机在江湖里选择一个人进行攻击，如果这个人在保护中，或者是出家人，官府中人，还有等级<21级的，将停止攻击，否则被攻击的人将人失去体力500,内力200,攻击NPC后你将是它的敌人,主人不攻击你自己的NPC,NPC有敌人后，那么别人是攻击不了的，只有它的敌人才能打!NPC被唤醒后，体力等状态会恢复如初!<br>
</font><strong>3.本江湖NPC具体参数!</strong><font color="#FFFFFF"><br>
同时最多召唤的NPC数量：<strong><font color="#00FF00"><%=application("aqjh_maxnpc")%></font></strong>个<br>
NPC物品暴出比例见上面：<strong><font color="#00FF00"><%=application("aqjh_npcwp")%></font></strong>值<br>
管理多少级的可以手动发放及踢出NPC：<strong><font color="#00FF00"><%=application("aqjh_npcff")%></font></strong>级<br>
<br>
</font>最后如果您对江湖NPC有什么建议、想法、要求请给我们提出！<br>
论坛地址：<a href=http://bbs.7758530.com target=_blank><font color=yellow>http://bbs.7758530.com</font></a><br>联系QQ:865240608</p>
</td>
</tr>
</table>
</body>
</html>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<%
'聊天室信息
Sub showchat(mess)
says=mess   '聊天数据
act="消息"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway&","&nowinroom&");parent.f2.document.af.npc.value='"&Application("aqjh_npc")&"';parent.m.location.reload();<"&"/script>"
addmsg saystr
end sub
Function Yushu1(a)
 Yushu1=(a and 31)
End Function
Sub AddMsg(Str)
Application.Lock()
Application("SayCount")=Application("SayCount")+1
i="SayStr"&YuShu1(Application("SayCount"))
Application(i)=Str
Application.UnLock()
End Sub
%>