<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../mywp.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"hcjs/jhjs")=0 then 
Response.Write "<script language=javascript>{alert('对不起，程序拒绝您的操作！！！\n     按确定返回！');parent.history.go(-1);}</script>" 
Response.End 
end if
yaoid=trim(request("id"))
if not isnumeric(yaoid) then Response.Redirect "../../error.asp?id=54"
yaopsl=int(abs(Request.form("sl")))
if yaopsl<1 or yaopsl>99 then
	Response.Redirect "../../error.asp?id=72"
end if
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(session("nowinroom")),"|")
fjname=chatinfo(0)
erase sjjh_roominfo
erase chatinfo
if fjname="高手房间" and Weekday(date())=7 and (Hour(time())>=20 and Hour(time())<=22) then
	Response.Write "<script language=javascript>{alert('夺宝期间不许购买暗器及药品！');parent.history.go(-1);}</script>" 
	Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT a,b,h,l,m,n FROM b where ID=" & yaoid,conn,2,2
if rs("b")<>"毒药" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../../error.asp?id=72"
	response.end
end if
wpname=rs("a")
tj=rs("l")
yin=rs("h")
jinbi=rs("m")
rs.close
rs.open "select 会员等级,银两,操作时间,w2,金币 from 用户 where 姓名='"&sjjh_name&"' and "&tj,conn,3,3
If Rs.Bof OR Rs.Eof then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你不具备购买条件\n["&replace(tj,chr(39),"")&"]！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
hy=rs("会员等级")
if hy>0 then
	yin=int(yin/2)
end if
sj=DateDiff("s",rs("操作时间"),now())
if sj<10 then
	s=10-sj
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('对不起系统限制，请等["&s&"秒钟]再操作！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if	
yin=yin*yaopsl
if yin>rs("银两") or jinbi>rs("金币") then
	Response.Write "<script Language=Javascript>alert('提示：购买不成功，原因：你的银两不够了、或金币不够');location.href = 'javascript:history.go(-1)';</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
zstemp=add(rs("w2"),wpname,yaopsl)
conn.execute "update 用户 set 银两=银两-" & yin & ",金币=金币-"&jinbi&",操作时间=now(),w2='"&zstemp&"' where 姓名='"&sjjh_name&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<body topmargin="6" leftmargin="0" bgcolor="#FFFFFF" background="../JHZB/bg.gif">
<%
if hy>0 then
	Response.Write "<script Language=Javascript>alert('提示：会员购买物品打5折,购买"&wpname&yaopsl&"个成功！');location.href = 'javascript:history.go(-1)';</script>"
else
	Response.Write "<script Language=Javascript>alert('提示：购买"&wpname&yaopsl&"个成功,如果您是会员可以打5折！');location.href = 'javascript:history.go(-1)';</script>"
end if
%>
