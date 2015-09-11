<!--#include file="../showchat.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Response.Expires=0
aqjh_name=Session("aqjh_name")
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能进行操作，进行此操作必须进入聊天室！');window.close();</script>"
	Response.End
end if
ndate=Day(date())
if ndate<27 then
		Response.Write "<script Language=Javascript>alert('提示：现在还没有到领奖的时间！');window.close();</script>"
		response.end
end if
myok=trim(lcase(cstr(server.HTMLEncode(Request("ok")))))
%>
<html><title>江湖秀大赛报名</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="style.css" rel=stylesheet></head>
<body leftmargin="0" topmargin="0">
<%act=clng(Request("act"))
if act<>1 and act<>2 and act<>3 and act<>4 then
	Response.Write "<script Language=Javascript>alert('提示：参数不正确！');window.close();</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("showmdb")
rs.Open "select * from pool", conn, 1,3
select case act
case 1
	mess="男子组第一名"
	if trim(rs("p_男第一名"))<>Session("aqjh_name")  then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script Language=Javascript>alert('提示：你搞错了,你并没有得奖！');window.close();</script>"
		response.end
	end if
	if rs("p_男第一领奖")=true then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script Language=Javascript>alert('提示：你已经领过奖了,不能再操作！');window.close();</script>"
		response.end
	end if
	rs("p_男第一领奖")=true
case 2
	mess="男子组第二名"
	if trim(rs("p_男第二名"))<>Session("aqjh_name")  then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script Language=Javascript>alert('提示：你搞错了,你并没有得奖！');window.close();</script>"
		response.end
	end if
	if rs("p_男第二领奖")=true then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script Language=Javascript>alert('提示：你已经领过奖了,不能再操作！');window.close();</script>"
		response.end
	end if
	rs("p_男第二领奖")=true
case 3
	mess="女子组第一名"
	if trim(rs("p_女第一名"))<>Session("aqjh_name")  then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script Language=Javascript>alert('提示：你搞错了,你并没有得奖！');window.close();</script>"
		response.end
	end if
	if rs("p_女第一领奖")=true then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script Language=Javascript>alert('提示：你已经领过奖了,不能再操作！');window.close();</script>"
		response.end
	end if
	rs("p_女第一领奖")=true
case 4
	mess="女子组第二名"
	if trim(rs("p_女第二名"))<>Session("aqjh_name")  then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script Language=Javascript>alert('提示：你搞错了,你并没有得奖！');window.close();</script>"
		response.end
	end if
	if rs("p_女第二领奖")=true then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script Language=Javascript>alert('提示：你已经领过奖了,不能再操作！');window.close();</script>"
		response.end
	end if
	rs("p_女第二领奖")=true
case else
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script Language=Javascript>alert('提示：参数不正确不能操作！');window.close();</script>"
		response.end
end select
'取出所有金币数
money=rs("p_金币")
'如果是第一名为30%
if act=1 or act=3 then
	gold=int(money*0.2)
else
	gold=int(money*0.1)
end if
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing
'打开用户库,增加金币
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 金币 from [用户] where 姓名='" & Session("aqjh_name") & "'",conn,1,3
rs("金币")=rs("金币")+gold
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=#ff0000>【消息】</font>["&session("aqjh_name")&"]<font color=red>在江湖大赛中获得"& mess &",得到将品金币<font class=t6>"& gold &"</font>个,新的江湖大赛将于下月1号举行,希望玩家作好准备!<font class=t>(" & time() & ")</font></font>"
call showchat(says)
Response.Write "<script Language=Javascript>alert('提示：将口今币领取成功,得到["& gold &"]个金币！');window.close();</script>"
response.end
%>
</body></html>