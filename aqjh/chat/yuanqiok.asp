<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc/func.asp"-->
<!--#include file="../config.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
if chatinfo(5)<>0 then
 Response.Write "<script language=javascript>{alert(''''提示：在["&chatinfo(0)&"]房间不可以发招！'''');}</script>"
 Response.End
end if
f=Minute(time()) 
%>
<div align="center"> 
  <p><font color="#FFFFFF"><span style='font-size:9pt'><font size="3">〖</font></span></font><a href="yuanqi.asp" target="f3"><font size="3" color="red">元气攻击</font></a><font size="3" color="#FFFFFF">〗</font>
  </p>
</div>
<%
chatbgcolor=Session("afa_chatbgcolor")
chatimage=Application("aqjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"chat")=0 then 
 Response.Write "<script language=javascript>{alert(''''对不起,程序拒绝您的操作！！！\n     按确定返回！'''');}</script>" 
 Response.End 
end if 
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
 Response.Write "<script Language=javascript>alert(''''你不能进行操作,进行此操作必须进入聊天室！'''');</script>"
 Response.End
end if
to1=trim(request.form("to1"))
hyjzzs=trim(request.form("hyjzzs"))
yuanqi=int(abs(clng(request.form("yuanqi"))))
zsdj=int(abs(clng(request.form("dj"))))
if Instr(dgjjzs,chr(39))<>0 or Instr(dgjjzs,chr(34))<>0 then
	Response.Write "<script Language=Javascript>alert('你要作什么滚远点！');location.href = 'javascript:history.go(-1)'</script>"
	response.end
end if
if Instr(LCase(application("aqjh_zanli")),LCase("!"&aqjh_name&"!"))>0  then
	Response.Write "<script Language=Javascript>alert('您正在“暂离”状态中,请使用“我回来啦”功能解除！');location.href = 'javascript:history.go(-1)'</script>"
	response.end
end if
if Instr(LCase(application("aqjh_zanli")),LCase("!"&to1&"!"))>0  then
	Response.Write "<script Language=Javascript>alert('"&to1&"正在“暂离”状态中,请不要攻击！');location.href = 'javascript:history.go(-1)'</script>"
	response.end
end if
if yuanqi<maxyuanqi then
	Response.Write "<script Language=Javascript>alert('提示：要使用["&dgjjzs&"],最少要使用"&maxyuanqi&"元气。');window.location.href='yuanqi.asp';</script>"
	Response.End
end if
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(to1)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('提示：所攻击的人不在江湖中！');window.location.href='yuanqi.asp';</script>"
	Response.End
end if
if to1="大家" or to1=Application("aqjh_automanname") then
	Response.Write "<script Language=Javascript>alert('提示：你不可以对大家或江湖机器人发招。');window.location.href='yuanqi.asp';</script>"
	Response.End
end if
if InStr(";" & Application("aqjh_npc"), ";" & to1 & "|")<>0 then
	Response.Write "<script Language=Javascript>alert('你不能对NPC使用此操作！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
stunt=""
stunt1=""
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from 用户 where 姓名='" & aqjh_name & "'",conn,2,2
if rs("门派")="出家" and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你是出家人不可以操作！');window.location.href='yuanqi.asp';}</script>"
	Response.End
end if
if rs("会员等级")<2 and aqjh_name<>"官府" then
	Response.Write "<script Language=Javascript>alert('提示：你的会员等级不够[2]级,不可以使用此招！');window.location.href='yuanqi.asp';</script>"
	myclose()
end if
sj=DateDiff("n",rs("死亡时间"),now())
if sj<5 and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你刚刚被别人杀死,还是先练练吧！');window.location.href='yuanqi.asp';}</script>"
	Response.End
end if
hy=rs("会员等级")
pdhy=rs("会员")
if hy=0 and pdhy=False then
 rs.close
 set rs=nothing
 conn.close
 set conn=nothing
 Response.Write "<script Language=javascript>alert('天呀,["&aqjh_name &"]你可不是会员,不能使用元气攻击！');window.close();</script>"
 response.end
end if
mp=rs("门派")
if rs("grade")>=6 and rs("grade")<10 then
	Response.Write "<script Language=Javascript>alert('提示：你是官府人员,不能使用["&hyjzzs&"]！');window.location.href='yuanqi.asp';</script>"
	myclose()
end if

if rs("元气")<yuanqi then
	Response.Write "<script Language=Javascript>alert('提示：你元气不足,不能发此招！');window.location.href='yuanqi.asp';</script>"
	myclose()
end if
if rs("保护")=True and aqjh_grade<10 then
	Response.Write "<script Language=Javascript>alert('提示：你正在练功保护中,不可以打架！');window.location.href='yuanqi.asp';</script>"
	myclose()
end if
if rs("杀人数")>=int(Application("aqjh_killman")) then
	Response.Write "<script Language=Javascript>alert('提示：你今天杀的人够多了,还想再杀人吗！');window.location.href='yuanqi.asp';</script>"
	myclose()
end if
sj=DateDiff("n",rs("操作时间"),now())
if sj<1 then
	s=1-sj
	Response.Write "<script Language=Javascript>alert('警告：请等"& s &"分再发招,可别累着！');window.location.href='yuanqi.asp';</script>"
	myclose()
end if
lxjwg1=rs("武功")
lxjgj1=rs("攻击")
mydj=rs("等级")
nn=int(mydj/10)+1
yuanqisx=rs("等级")*aqjh_yuanqisx+2000+rs("元气加")
yuanqibf=int(yuanqi/yuanqisx)
rs.close
rs.open "select * from 用户 where 姓名='" & to1 & "'",conn,2,2
zstt=rs("转生")
if zstt>=10 and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：对方已转生10次了,这招伤不了对方！');window.location.href='yuanqi.asp';}</script>"
	Response.End
end if
if rs("门派")="出家" and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：他是出家人不可以操作！');;window.location.href='yuanqi.asp';}</script>"
	Response.End
end if
sj=DateDiff("n",rs("死亡时间"),now())
if sj<8 and rs("宝物")="无" and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：他刚刚被别人杀死,还是先放过他吧！');window.location.href='yuanqi.asp';}</script>"
	Response.End
end if
jhhy=rs("会员等级")
ntnt=rs("等级")
tjf=rs("通缉")
if rs("grade")>=6 then
	Response.Write "<script Language=Javascript>alert('提示：他是官府人员,对他有什么意见你可以向站长反应！');window.location.href='yuanqi.asp';</script>"
	myclose()
end if
if rs("等级")<70 then
	Response.Write "<script Language=Javascript>alert('提示：他还不够70级,你不要这么狠！');window.location.href='yuanqi.asp';</script>"
	myclose()
end if
'if rs("保护")=true and hyjzzs<>"终极无限" then
tomp=rs("门派")
lgbh=rs("保护")
lxjwg2=rs("武功")
lxjgj2=rs("防御")
if hyjzzs="终极无限" then
yiyuanqiiang2=rs("银两")
else
cdyuanqi=rs("银两")
if cdyuanqi>=yuanqi*100 then
yiyuanqiiang1=yuanqi*100
else
yiyuanqiiang1=cdyuanqi
end if
end if
youid=rs("id")
randomize timer
sss=int(rnd*9)+1
qq1=lxjwg1+lxjgj1+100000-lxjwg2+lxjgj2
qq2=(tl+yuanqi)*sss
qq3=(qq1+qq2)*bf
qq4=sqr(qq3)
conn.execute "update 用户 set 道德=道德-100,元气=元气-" & yuanqi & ",操作时间=now() where 姓名='" & aqjh_name & "'"
e=""
qi2=conn.execute("select 元气 from 用户 where  姓名='"&aqjh_name&"'")
yyuanqi=qi2(0)
if jhhy<3 then
	conn.execute "update 用户 set 杀人数=杀人数+1,总杀人=总杀人+1 where 姓名='" & aqjh_name & "'"
	if rs("宝物")=Application("aqjh_baowuname") then
		conn.execute "update 用户 set 宝物修练=0,宝物='"& Application("aqjh_baowuname") &"' where id="&aqjh_id
		conn.execute "update 用户 set 宝物修练=0,宝物='无',内力=100,体力=2000 where 姓名='"& to1 &"'"
		stunt=aqjh_name & "把"& to1 &"的宝物:"& Application("aqjh_baowuname") &"抢走。江湖宝物需要进行修练才可以得到更多的东西！"
	else
	if hyjzzs="终极无限" then
		if lgbh=true then
			conn.execute "update 用户 set 保护=false,状态='死',银两=0,操作时间=now(),元气=0 where 姓名='" & to1 & "'"
			lgbh=false
			e="看来他要倒大霉了。"
			stunt=aqjh_name & "运足" & yuanqi & "点元气,对<img src='img/41.gif'><font color=blue>" & to1 & "</font>使用元气攻击之<font color=008000>【"&hyjzzs&"】</font>,去除了<font color=blue>" & to1 & "</font>的<font color=red>练功保护<font/>。" & e
		else
			conn.execute "update 用户 set 状态='死',银两=0,元气=0 where 姓名='" & to1 & "'"
			conn.execute "update 用户 set allvalue=allvalue+50,银两=银两+" & yiyuanqiiang2 & " where 姓名='" & aqjh_name & "'"
			
			e="点," & to1 & "慢慢的<img src=xx/gif/WG8.GIF>倒了下去……从此江湖上又少了一只大虾," & aqjh_name & "得到" & to1 & "的所有现金" & yiyuanqiiang2 & "两,存点[50]点,现有元气"&yyuanqi&"点," & to1 & "的所有物品从此散落江湖!"
			fn1=hyjzzs
			call boot(to1,"元气攻击,操作者："&aqjh_name&",["&mp&"]"&fn1)
			conn.execute "insert into l(b,a,c,e,d) values ('" & to1 & "',now(),'" & aqjh_name & "','元气攻击之"&hyjzzs&"','人命')"
		end if
	else
if lgbh=true then
	conn.execute "update 用户 set 道德=道德+100,元气=元气+" & yuanqi & ",操作时间=now() where 姓名='" & aqjh_name & "'"
    Response.Write "<script Language=Javascript>alert('他正在练功,你的攻击力无法破除他的练功保护,还是先练练再说吧！');window.location.href='javascript:history.go(-1)';</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.End
else
		conn.execute "update 用户 set 状态='死',银两=0,事件原因='"&aqjh_name&"|元气攻击之"&hyjzzs&"' where 姓名='" & to1 & "'"
		if tjf=True then
			conn.execute "update 用户 set 银两=0,存款=int(存款/2),道德=0,魅力=0 where 姓名='" & to1 & "'"
		end if
		conn.execute "update 用户 set allvalue=allvalue+500,银两=银两+" & yiyuanqiiang1 & " where 姓名='" & aqjh_name & "'"
		e="点," & to1 & "慢慢的<img src=xx/gif/WG8.GIF>倒了下去……从此江湖上又少一个讨厌鬼," & aqjh_name & "得到" & to1 & "的现金" & yiyuanqiiang1 & "两！存点[500]点！现有元气"&yyuanqi&"点"
		fn1=hyjzzs
		call boot(to1,"元气攻击,操作者："&aqjh_name&",["&mp&"]"&fn1)
		conn.execute "insert into l(b,a,c,e,d) values ('" & to1 & "',now(),'" & aqjh_name & "','元气攻击之"&hyjzzs&"','人命')"
	end if
	end if
end if
else
 conn.execute("update 用户 set 元气=元气-"&yuanqi&",操作时间=now() where 姓名='"&to1&"'")
end if
if e="" then
	e="点。"
end if
if stunt="" then
	stunt=aqjh_name & "运足" & yuanqi & "点元气,对<img src='img/021.gif'><font color=blue>" & to1 & "</font>使用元气攻击之<font color=008000>【"&hyjzzs&"】</font>,杀伤" & yuanqi & " "& e
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=red><b>【元气攻击"&hyjzzs&"】</b>"& stunt &"</font>"& t
says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
act="消息"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& session("nowinroom") & ");<"&"/script>"
addmsg saystr
Function Yushu(a)
	Yushu=(a and 31)
End Function
Sub AddMsg(Str)
Application.Lock()
Application("SayCount")=Application("SayCount")+1
i="SayStr"&YuShu(Application("SayCount"))
Application(i)=Str
Application.Unlock()
End Sub
Response.Write "<script Language=Javascript>alert('恭喜,您的"&hyjzzs&"已经发招完成！');window.location.href='yuanqi.asp';</script>"
function myclose()
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.End
end function
%>
