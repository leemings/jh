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
  <p><font color="#FFFFFF"><span style='font-size:9pt'><font size="3">〖</font></span></font><a href="zjjj.asp" target="f3"><font size="3" color="#FF0000">终极绝技</font></a><font size="3" color="#FFFFFF">〗</font>
  </p>
</div>
<%
chatbgcolor=Session("afa_chatbgcolor")
chatimage=Application("aqjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"chat")=0 then 
 Response.Write "<script language=javascript>{alert(''''对不起，程序拒绝您的操作！！！\n     按确定返回！'''');}</script>" 
 Response.End 
end if 
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
 Response.Write "<script Language=javascript>alert(''''你不能进行操作，进行此操作必须进入聊天室！'''');</script>"
 Response.End
end if
to1=trim(request.form("to1"))
hyjzzs=trim(request.form("hyjzzs"))
tl=int(abs(clng(request.form("tl"))))
nl=int(abs(clng(request.form("nl"))))
maxtl=int(abs(clng(request.form("maxtl"))))
maxnl=int(abs(clng(request.form("maxnl"))))
money=int(abs(clng(request.form("money"))))
zsdj=int(abs(clng(request.form("dj"))))
if Instr(dgjjzs,chr(39))<>0 or Instr(dgjjzs,chr(34))<>0 then
	Response.Write "<script Language=Javascript>alert('你要作什么滚远点！');location.href = 'javascript:history.go(-1)'</script>"
	response.end
end if
if Instr(LCase(application("aqjh_zanli")),LCase("!"&aqjh_name&"!"))>0  then
	Response.Write "<script Language=Javascript>alert('您正在“暂离”状态中，请使用“我回来啦”功能解除！');location.href = 'javascript:history.go(-1)'</script>"
	response.end
end if
if Instr(LCase(application("aqjh_zanli")),LCase("!"&to1&"!"))>0  then
	Response.Write "<script Language=Javascript>alert('"&to1&"正在“暂离”状态中，请不要攻击！');location.href = 'javascript:history.go(-1)'</script>"
	response.end
end if
if tl<maxtl or nl<maxnl then
	Response.Write "<script Language=Javascript>alert('提示：要使用["&dgjjzs&"]，最少要使用"&maxtl&"体力，"&maxnl&"内力。');</script>"
	Response.End
end if
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(to1)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('提示：所攻击的人不在江湖中！');</script>"
	Response.End
end if
if to1="大家" or to1=Application("aqjh_automanname") then
	Response.Write "<script Language=Javascript>alert('提示：你不可以对大家或江湖机器人发招。');</script>"
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
sj=DateDiff("n",rs("死亡时间"),now())
if sj<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你刚刚被别人杀死，还是先练练吧！');}</script>"
	Response.End
end if
hy=rs("会员等级")
pdhy=rs("会员")
if hy<2 and pdhy=False then
 rs.close
 set rs=nothing
 conn.close
 set conn=nothing
 Response.Write "<script Language=javascript>alert('天呀，["&aqjh_name &"]你可不是2级会员,不能使用此特技！');window.close();</script>"
 response.end
end if
if rs("银两")<money then
	Response.Write "<script Language=Javascript>alert('提示：你没有"&money&"两银子，不能使用["&hyjzzs&"]！');</script>"
	myclose()
end if

if rs("内力")<nl or rs("体力")<tl then
	Response.Write "<script Language=Javascript>alert('提示：你内力或体力不足，不能发此招！');</script>"
	myclose()
end if
if rs("保护")=True then
	Response.Write "<script Language=Javascript>alert('提示：你正在练功保护中，不可以打架！');</script>"
	myclose()
end if
if rs("等级")<90 then
	Response.Write "<script Language=Javascript>alert('提示：你的等级不够[90],不可以使用此招！');</script>"
	myclose()
end if
if rs("智力")<50 then
	Response.Write "<script Language=Javascript>alert('提示：此招需要50点智力！');</script>"
	myclose()
end if
if rs("知质")<50 then
	Response.Write "<script Language=Javascript>alert('提示：此招需要50点知质！');</script>"
	myclose()
end if
if rs("泡豆点数")<100 then
	Response.Write "<script Language=Javascript>alert('提示：此招需要100点豆点！');</script>"
	myclose()
end if
if rs("杀人数")>=int(Application("aqjh_killman")) then
	Response.Write "<script Language=Javascript>alert('提示：你今天杀的人够多了，还想再杀人吗！');</script>"
	myclose()
end if
sj=DateDiff("n",rs("操作时间"),now())
if sj<20 and instr(Application("aqjh_user"),aqjh_name)=0 then
	s=20-sj
	Response.Write "<script Language=Javascript>alert('警告：请等"& s &"分再发招,可别累着！');</script>"
	myclose()
end if
lxjwg1=rs("武功")
lxjgj1=rs("攻击")
mydj=rs("等级")
nn=int(mydj/10)+1
tlsx=rs("等级")*aqjh_tlsx+5260+rs("体力加")
nlsx=rs("等级")*aqjh_nlsx+2000+rs("内力加")
wgsx=rs("等级")*aqjh_wgsx+3800+rs("武功加")
tlbf=int(tl/tlsx)
nlbf=int(nl/nlsx)
wgbf=int(lxjwg1/wgsx)
bf=int((nlbf+tlbf+wgbf)/3)
rs.close
rs.open "select * from 用户 where 姓名='" & to1 & "'",conn,2,2
sj=DateDiff("n",rs("死亡时间"),now())
if sj<10 and rs("宝物")="无" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：他刚刚被别人杀死，还是先放过他吧！');}</script>"
	Response.End
end if
jhhy=rs("会员等级")
ntnt=rs("等级")
tjf=rs("通缉")
if rs("等级")<90 then
	Response.Write "<script Language=Javascript>alert('提示：他还不够90级，你不要这么狠！');</script>"
	myclose()
end if
if rs("体力")<10000 and hyjzzs<>"冰之封舞" then
	Response.Write "<script Language=Javascript>alert('提示：他体力小于10000,除非使用特别招式方可杀他！');</script>"
	myclose()
end if
if rs("保护")=true and hyjzzs<>"终极杀着" and instr(Application("aqjh_user"),aqjh_name)=0 then
	Response.Write "<script Language=Javascript>alert('他已正在练功，你的攻击无效，以后再说吧！');location.href = 'javascript:history.go(-1)';</script>"
	myclose()
end if
tomp=rs("门派")
sf=rs("身份")
lgbh=rs("保护")
lxjwg2=rs("武功")
lxjgj2=rs("防御")
yinliang2=rs("银两")
youid=rs("id")
randomize timer
sss=int(rnd*9)+1
qq1=lxjwg1+lxjgj1+100000-lxjwg2+lxjgj2
qq2=(tl+nl)*sss
qq3=(qq1+qq2)*bf
qq4=sqr(qq3)
killer=int((qq4*nn)/2)
if killer<=100 then
	randomize timer
	killer=int(rnd*99)+1
end if
killer=killer+(money/50)
randomize timer
m1 = Int(100 * Rnd)+100
gjtl=int(fn1/m1)
killer=killer+gjtl+(tl+nl)/3+money/80
shengtili=rs("体力")-killer
conn.execute "update 用户 set 体力=体力-"  & killer & " where 姓名='" & to1 & "'"
conn.execute "update 用户 set 知质=知质-50,智力=智力-50,泡豆点数=泡豆点数-100,道德=道德-100,银两=银两-"&money&",体力=体力-" & tl & ",内力=内力-" & nl & ",操作时间=now() where 姓名='" & aqjh_name & "'"
e=""
if shengtili<-100 then
	conn.execute "update 用户 set 杀人数=杀人数+1,总杀人=总杀人+1 where 姓名='" & aqjh_name & "'"
	if rs("宝物")=Application("aqjh_baowuname") then
		conn.execute "update 用户 set 宝物修练=0,宝物='"& Application("aqjh_baowuname") &"' where id="&aqjh_id
		conn.execute "update 用户 set 宝物修练=0,宝物='无',内力=100,体力=2000 where 姓名='"& to1 &"'"
		stunt=aqjh_name & "把"& to1 &"的宝物:"& Application("aqjh_baowuname") &"抢走。江湖宝物需要进行修练才可以得到更多的东西！"
	else
	if hyjzzs="终极杀着" then
		if lgbh=true then
			conn.execute "update 用户 set 保护=false,体力=体力+" & killer & ",操作时间=now() where 姓名='" & to1 & "'"
			lgbh=false
			e="看来他命不久了。"
			stunt=aqjh_name & "为[" &tomp& "]的[" &sf& "],运足" & nl & "点内力，" & tl & "点体力，对<img src='img/41.gif'><font color=blue>" & to1 & "</font>使用了江湖无界的终极绝技之<font color=008000>【"&hyjzzs&"】</font>，去除了<font color=blue>" & to1 & "</font>的<font color=red>练功保护<font/>。" & e
		else
			conn.execute "update 用户 set 状态='死',事件原因='"&aqjh_name&"|终极绝技之"&hyjzzs&"' where 姓名='" & to1 & "'"
			conn.execute "update 用户 set allvalue=allvalue+50 where 姓名='" & aqjh_name & "'"
			
			e="点，" & to1 & "<bgsound src=wav/si.wav loop=1>慢慢的<img src=xx/gif/WG8.GIF>倒了下去……从此快乐江湖上又少了一只大虾," & aqjh_name & "得到了系统奖励的存点[50]点!"
			fn1=hyjzzs
			call boot(to1,"终极绝技，操作者："&aqjh_name&",["&mp&"]"&fn1)
			conn.execute "insert into l(b,a,c,e,d) values ('" & to1 & "',now(),'" & aqjh_name & "','终极绝技之"&hyjzzs&"','人命')"
		end if
	else
		conn.execute "update 用户 set 状态='死',事件原因='"&aqjh_name&"|终极绝技之"&hyjzzs&"' where 姓名='" & to1 & "'"
		if tjf=True then
			conn.execute "update 用户 set 银两=0,存款=int(存款/2),道德=0,魅力=0 where 姓名='" & to1 & "'"
		end if
		conn.execute "update 用户 set allvalue=allvalue+500,智力=智力+50 where 姓名='" & aqjh_name & "'"
		e="点，" & to1 & "<bgsound src=wav/si.wav loop=1>慢慢的<img src=xx/gif/WG8.GIF>倒了下去……从此江湖上又少了一只大虾," & aqjh_name & "得到了系统奖励的智力50点！存点[500]点！"
		fn1=hyjzzs
		call boot(to1,"终极绝技，操作者："&aqjh_name&"，["&mp&"]"&fn1)
		conn.execute "insert into l(b,a,c,e,d) values ('" & to1 & "',now(),'" & aqjh_name & "','终极绝技之"&hyjzzs&"','人命')"
	end if
	end if
end if
if lgbh=true then
	conn.execute "update 用户 set 道德=道德+100,银两=银两+11002000,体力=体力+" & tl & ",内力=内力+" & nl & ",操作时间=now() where 姓名='" & aqjh_name & "'"
    Response.Write "<script Language=Javascript>alert('他已正在练功，你的攻击力无法破除他的练功保护，还是先练练再说吧！');</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.End
end if
if e="" then
	e="点。"
end if
if stunt="" then
	stunt=aqjh_name & "<bgsound src=wav/dgjj0.wav loop=2>运足" & nl & "点内力，" & tl & "点体力，对<img src='img/021.gif'><font color=blue>" & to1 & "</font>使用了江湖上久已传的终极绝技之<font color=008000>【"&hyjzzs&"】</font>，杀伤" & killer & e
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=red><b>【终极绝技"&hyjzzs&"】</b>"& stunt &"</font>"& t
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
Application.UnLock()
End Sub
Response.Write "<script Language=Javascript>alert('恭喜，您的"&hyjzzs&"已经发招完成！');</script>"
function myclose()
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.End
end function
%>