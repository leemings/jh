<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../const.asp"-->
<!--#include file="../wzconfig.asp"-->
<!--#include file="../config.asp"-->
<body leftmargin="0" topmargin="0" bgcolor="#CCCCCC" background="../jhimg/bgcheetah.gif">
<%
call chkpost()
if Instr(LCase(Application("axjh_useronlinename"&nowinroom))," "&LCase(axjh_name)&" ")=0 or session("axjh_inthechat")<>"1" then
	Response.Write "<script Language=Javascript>alert('你不能进行操作，进行此操作必须进入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
axjh_roominfo=split(Application("axjh_room"),";")
chatinfo=split(axjh_roominfo(nowinroom),"|")
fjname=chatinfo(0)
erase chatinfo
erase axjh_roominfo
if fjname=application("axjh_dbroom") then
	erase axjh_roominfo
	erase chatinfo
	Response.Write "<script language=JavaScript>{alert('提示：夺宝大赛期间不可以使用密室练武！');location.href = 'javascript:history.back()';}</script>"
	Response.End
end if
money=int(abs(request.form("money")))
if money<>1000 and  money<>10000 and money<>100000 and money<>150000 and money<>200000 and money<>1000000 then 
	Response.Write "<script Language=Javascript>alert('数据出错');location.href = 'javascript:history.back()';</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("axjh_usermdb")
rs.open "select 等级,体力,内力,武功,武功加,会员等级,会员,操作时间 from 用户 where 姓名='"&axjh_name&"'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../error.asp?id=210"
	response.end
end if
hydj=rs("会员等级")
select case hydj
	case 0
		bf=1.2
		jgsj=180
	case 1
		bf=1.25
		jgsj=180
	case 2
		bf=1.3
		jgsj=120
	case 3
		bf=1.35
		jgsj=60
	case 4
		bf=1.4
		jgsj=0
end select
sj=DateDiff("s",rs("操作时间"),now())
if sj<jgsj then
	s=jgsj-sj
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你是"&hydj&"级会员，"&jgsj&"秒一次，请等["&s&"秒]！');location.href = 'javascript:history.back()';}</script>"
	Response.End
end if
mytl=rs("体力")
mynl=rs("内力")
wg=rs("武功")
wgj=rs("武功加")
pdhy=rs("会员")
zddj=rs("等级")
rs.close
if pdhy<>false then
	bf=1.3
end if
wgsx=int((zddj*axjh_wgsx+3800+wgj)*bf)
xtl=money*10
xnl=money*5
newwg=wg+money
if wg>=wgsx then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你的武功已达到上限的"& bf &"倍了，升级后再来练吧！');location.href = 'javascript:history.back()';</script>"
	response.end
end if
if xtl>mytl or xnl>mynl then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('修练此种武功可增加"&money&"点武功，但需消耗体力："&xtl&"，内力："&xnl&"，你目前体力值为："&mytl&"，内力值为："&mynl&"，不足以修练。');location.href = 'javascript:history.back()';</script>"
	response.end
end if
if newwg>=wgsx then
	conn.execute "update 用户 set 武功="&wgsx&",体力=体力-"& xtl &",内力=内力-"& xnl &",操作时间=now() where 姓名='"&axjh_name&"'"
	set rs=nothing
	conn.close
	set conn=nothing
	ltsay="<b><font color=#FF0000>【闭关练武】</font></b><font color=blue>"&axjh_name&"</font>在密室闭关修习，武功大进，武功上涨<b><font color=#FF0000>+"&money&"</font></b>点，消耗<font color=#0000FF>体力</font>：<b><font color=#FF0000>-"&xtl&"</font></b>点，<font color=#0000FF>内力</font><font color=#FF0000><b>-"&xnl&"</b></font>点，武功已达目前状态最高境界，各位要小心了。"
	call lts("消息","大家",ltsay,0,nowinroom)
	Response.Write "<script Language=Javascript>alert('你通过闭关修练！终于使自己的武功大进,达到当前最高境界，武功上涨："&money&"点，现已达到上限的"& bf &"倍，升级后再来练吧！');window.close();</script>"
else
	conn.execute "update 用户 set 武功=武功+"& money &",体力=体力-"& xtl &",内力=内力-"& xnl &",操作时间=now() where 姓名='"&axjh_name&"'"
	conn.close
	set rs=nothing
	set conn=nothing
	ltsay="<b><font color=#FF0000>【闭关练武】</font></b><font color=blue>"&axjh_name&"</font>在密室闭关修习，武功大进，武功上涨<b><font color=#FF0000>+"&money&"</font></b>点，消耗<font color=#0000FF>体力</font>：<b><font color=#FF0000>-"&xtl&"</font></b>点，<font color=#0000FF>内力</font><font color=#FF0000><b>-"&xnl&"</b></font>点。"
	call lts("消息","大家",ltsay,0,nowinroom)
	Response.Write "<script Language=Javascript>alert('你在密室潜心研习武学，终于使自己的武功大进,武功：+"& money &"点，消耗体力：-"&xtl&"点，内力：-"&xnl&"点！');location.href = 'javascript:history.back()';</script>"
end if
%>
